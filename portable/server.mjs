import http from 'node:http';
import https from 'node:https';
import { readFile, stat } from 'node:fs/promises';
import { resolve, sep, extname } from 'node:path';
import { fileURLToPath } from 'node:url';

const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.woff2': 'font/woff2',
};

export function createStaticHandler(root) {
  const base = resolve(root);
  return async (request, response) => {
    response.setHeader('X-Content-Type-Options', 'nosniff');
    response.setHeader('Cache-Control', 'no-store');
    if (!/^(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?$/i.test(request.headers.host ?? '')) {
      response.writeHead(403).end();
      return;
    }
    if (!['GET', 'HEAD'].includes(request.method)) {
      response.writeHead(405).end();
      return;
    }
    let path;
    try {
      path = decodeURIComponent((request.url ?? '/').split('?')[0]);
    } catch {
      response.writeHead(400).end();
      return;
    }
    if (path === '/health') {
      response.setHeader('Content-Type', 'application/json');
      response.end(JSON.stringify({ product: 'office-helper-portable', version: 1 }));
      return;
    }
    if (path.includes('\\') || path.includes('\0') || path.split('/').includes('..')) {
      response.writeHead(403).end();
      return;
    }
    if (path.endsWith('/')) path += 'index.html';
    if (!(path === '/index.html' || path === '/favicon.svg' || /^\/(assets|wps)\//.test(path))) {
      response.writeHead(404).end();
      return;
    }
    const filename = resolve(base, '.' + path);
    if (!filename.startsWith(base + sep) || !mime[extname(filename)]) {
      response.writeHead(404).end();
      return;
    }
    try {
      const info = await stat(filename);
      if (!info.isFile()) throw new Error('Not a file');
      const body = await readFile(filename);
      response.setHeader('Content-Type', mime[extname(filename)]);
      response.setHeader('Content-Length', body.length);
      response.end(request.method === 'HEAD' ? undefined : body);
    } catch {
      response.writeHead(404).end();
    }
  };
}

export async function startServers(config) {
  const servers = [];
  const handler = createStaticHandler(config.siteDirectory);
  async function listen(server, port, host) {
    servers.push(server);
    await new Promise((ok, fail) => {
      server.once('error', fail);
      server.listen(port, host, () => {
        server.removeListener('error', fail);
        ok();
      });
    });
  }
  try {
    // Bind only loopback. Office's localhost can resolve to either address family.
    await listen(http.createServer(handler), config.httpPort, '127.0.0.1');
    if (config.office) {
      const options = {
        cert: await readFile(config.certificatePath),
        key: await readFile(config.privateKeyPath),
        minVersion: 'TLSv1.2',
      };
      await listen(https.createServer(options, handler), config.httpsPort, '127.0.0.1');
      await listen(https.createServer(options, handler), config.httpsPort, '::1');
    }
  } catch (error) {
    for (const server of servers) {
      server.closeAllConnections();
      server.close();
    }
    throw error;
  }
  return async () => {
    await Promise.all(
      servers.map(
        (server) =>
          new Promise((done) => {
            server.closeAllConnections();
            server.close(() => done());
          }),
      ),
    );
  };
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const config = JSON.parse(await readFile(process.argv[2], 'utf8'));
    const stop = await startServers(config);
    console.log(
      JSON.stringify({
        ready: true,
        httpPort: config.httpPort,
        httpsPort: config.office ? config.httpsPort : null,
      }),
    );
    let stopping = false;
    const shutdown = async () => {
      if (stopping) return;
      stopping = true;
      await stop();
      process.exit(0);
    };
    process.stdin.resume();
    process.stdin.once('end', shutdown);
    process.stdin.on('data', (data) => {
      if (data.toString().trim() === 'stop') void shutdown();
    });
    process.once('SIGINT', shutdown);
    process.once('SIGTERM', shutdown);
    // No parent process means the launcher crashed or was terminated.
    const parent = process.ppid;
    setInterval(() => {
      try {
        process.kill(parent, 0);
      } catch {
        void shutdown();
      }
    }, 5000).unref();
  } catch (error) {
    console.error(
      JSON.stringify({
        error: error.code === 'EADDRINUSE' ? 'PORT_IN_USE' : 'START_FAILED',
        detail: error.message,
      }),
    );
    process.exitCode = 1;
  }
}
