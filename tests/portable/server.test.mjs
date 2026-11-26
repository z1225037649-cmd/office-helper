import { test } from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createStaticHandler, startServers } from '../../portable/server.mjs';

test('portable server serves built assets only and rejects hostile requests', async () => {
  const root = await mkdtemp(join(tmpdir(), 'office-helper-server-'));
  await mkdir(join(root, 'assets'));
  await writeFile(join(root, 'index.html'), '<h1>Helper</h1>');
  await writeFile(join(root, 'assets/app.js'), 'window.helper=true');
  await writeFile(join(root, 'private.key'), 'not-served');
  const server = http.createServer(createStaticHandler(root));
  await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
  const port = server.address().port;
  const request = (path, options = {}) =>
    new Promise((ok, fail) => {
      const req = http.request({ host: '127.0.0.1', port, path, ...options }, (response) => {
        let body = '';
        response.on('data', (data) => (body += data));
        response.on('end', () =>
          ok({ status: response.statusCode, body, type: response.headers['content-type'] }),
        );
      });
      req.on('error', fail);
      req.end();
    });
  try {
    assert.equal((await request('/?host=office')).body, '<h1>Helper</h1>');
    assert.equal((await request('/assets/app.js')).status, 200);
    assert.equal((await request('/private.key')).status, 404);
    assert.equal((await request('/native-check.html')).status, 404);
    assert.equal((await request('/src/main.tsx')).status, 404);
    assert.equal((await request('/assets/%2e%2e/private.key')).status, 403);
    assert.equal((await request('/assets/%5c..%5cprivate.key')).status, 403);
    assert.equal((await request('/%', {})).status, 400);
    assert.equal((await request('/', { method: 'POST' })).status, 405);
    assert.equal((await request('/', { headers: { Host: 'evil.example' } })).status, 403);
    assert.equal((await request('/', { method: 'HEAD' })).body, '');
  } finally {
    server.closeAllConnections();
    await new Promise((ok) => server.close(ok));
    await rm(root, { recursive: true, force: true });
  }
});

test('a port conflict leaves the existing service intact', async () => {
  const occupied = http.createServer((_req, res) => res.end('existing'));
  await new Promise((ok) => occupied.listen(0, '127.0.0.1', ok));
  const port = occupied.address().port;
  try {
    await assert.rejects(startServers({ siteDirectory: '.', httpPort: port, office: false }), {
      code: 'EADDRINUSE',
    });
    assert.equal(await (await fetch(`http://127.0.0.1:${port}`)).text(), 'existing');
  } finally {
    occupied.closeAllConnections();
    await new Promise((ok) => occupied.close(ok));
  }
});
