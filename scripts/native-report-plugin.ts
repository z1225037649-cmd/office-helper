import type { Plugin } from 'vite';
import { appendFile, mkdir, readFile } from 'node:fs/promises';

export function nativeReportPlugin(): Plugin {
  return {
    name: 'local-native-test-report',
    configureServer(server) {
      server.middlewares.use((request, _response, next) => {
        if (process.env.HELPER_NATIVE_TEST_TOKEN && request.url?.startsWith('/wps')) {
          const record =
            JSON.stringify({
              at: new Date().toISOString(),
              path: request.url.split('?')[0],
              userAgent: request.headers['user-agent'],
            }) + '\n';
          void mkdir('output/native', { recursive: true }).then(() =>
            appendFile('output/native/wps-requests.jsonl', record),
          );
        }
        next();
      });
      server.middlewares.use('/wps-native-test', (request, response, next) => {
        if (!process.env.HELPER_NATIVE_TEST_TOKEN) {
          next();
          return;
        }
        const relative = (request.url ?? '/').split('?')[0]!.replace(/^\/(writer|sheet)/, '');
        const file = relative === '/' ? 'index.html' : relative.slice(1);
        if (file === 'manifest.xml') {
          response.setHeader('Content-Type', 'application/xml');
          response.end(
            '<?xml version="1.0"?><JsPlugin><ApiVersion>1.0.0</ApiVersion><Name>' +
              (request.url?.includes('sheet') ? 'OfficeHelperSheet' : 'OfficeHelperWriter') +
              '</Name><Description>Native test</Description></JsPlugin>',
          );
          return;
        }
        if (!['index.html', 'main.js', 'ribbon.xml'].includes(file ?? '')) {
          response.statusCode = 404;
          response.end();
          return;
        }
        void readFile(`.local/wps-native-test/${file}`)
          .then((content) => {
            response.setHeader(
              'Content-Type',
              file?.endsWith('.js')
                ? 'application/javascript'
                : file?.endsWith('.xml')
                  ? 'application/xml'
                  : 'text/html',
            );
            response.end(content);
          })
          .catch(() => {
            response.statusCode = 404;
            response.end();
          });
      });
      server.middlewares.use('/__native-report', (request, response) => {
        if (
          !process.env.HELPER_NATIVE_TEST_TOKEN ||
          request.headers['x-helper-test'] !== process.env.HELPER_NATIVE_TEST_TOKEN ||
          request.method !== 'POST'
        ) {
          response.statusCode = 403;
          response.end();
          return;
        }
        let body = '';
        request.on('data', (chunk) => {
          body += chunk;
          if (body.length > 10000) request.destroy();
        });
        request.on('end', () => {
          void (async () => {
            try {
              const report = JSON.parse(body);
              await mkdir('output/native', { recursive: true });
              await appendFile(
                'output/native/results.jsonl',
                JSON.stringify({ at: new Date().toISOString(), ...report }) + '\n',
              );
              response.end('ok');
            } catch {
              response.statusCode = 400;
              response.end('invalid report');
            }
          })();
        });
      });
    },
  };
}
