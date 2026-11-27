import { test } from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { copyFile, mkdtemp, mkdir, writeFile, rm, symlink } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createStaticHandler, startServers } from '../../portable/server.mjs';

test('portable server serves built assets only and rejects hostile requests', async () => {
  const root = await mkdtemp(join(tmpdir(), 'office-helper-server-'));
  await mkdir(join(root, 'assets'));
  await writeFile(join(root, 'index.html'), '<h1>Helper</h1>');
  await writeFile(join(root, 'assets/app.js'), 'window.helper=true');
  await writeFile(join(root, 'private.key'), 'not-served');
  await writeFile(join(root, 'assets/private.key'), 'not-served');
  await writeFile(join(root, 'assets/native-check.js'), 'not-served');
  await writeFile(join(root, 'assets/app.js.map'), 'not-served');
  await mkdir(join(root, 'wps/writer'), { recursive: true });
  await writeFile(join(root, 'wps/writer/native-check.html'), 'not-served');
  await writeFile(join(root, 'wps/writer/index.html'), '<h1>WPS</h1>');
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
    assert.equal((await request('/assets/private.key')).status, 404);
    assert.equal((await request('/assets/native-check.js')).status, 404);
    assert.equal((await request('/assets/app.js.map')).status, 404);
    assert.equal((await request('/wps/writer/native-check.html')).status, 404);
    assert.equal((await request('/wps/writer/')).body, '<h1>WPS</h1>');
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

test('portable server refuses an assets junction pointing outside the site', async () => {
  const root = await mkdtemp(join(tmpdir(), 'office-helper-server-link-'));
  const site = join(root, 'site');
  const outside = join(root, 'outside');
  await mkdir(site);
  await mkdir(outside);
  await writeFile(join(outside, 'app.js'), 'outside-site-sentinel');
  await symlink(outside, join(site, 'assets'), process.platform === 'win32' ? 'junction' : 'dir');
  const server = http.createServer(createStaticHandler(site));
  await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
  try {
    const result = await fetch(`http://127.0.0.1:${server.address().port}/assets/app.js`);
    assert.equal(result.status, 404);
    assert.ok(!(await result.text()).includes('outside-site-sentinel'));
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

test(
  'portable CLI starts and stops from a Chinese path with no npm on PATH',
  { timeout: 15000 },
  async (t) => {
    const root = await mkdtemp(join(tmpdir(), 'office-helper-cli-中文 空格-'));
    let child;
    t.after(async () => {
      if (child?.pid && child.exitCode === null && child.signalCode === null) {
        const closed = once(child, 'close');
        child.kill();
        await closed;
      }
      await rm(root, { recursive: true, force: true });
    });
    await mkdir(join(root, 'app'));
    await mkdir(join(root, 'site'));
    for (const file of ['server.mjs', 'site-policy.mjs'])
      await copyFile(new URL(`../../portable/${file}`, import.meta.url), join(root, 'app', file));
    await writeFile(join(root, 'site/index.html'), '<h1>Isolated helper</h1>');
    const probe = http.createServer();
    await new Promise((ok) => probe.listen(0, '127.0.0.1', ok));
    const port = probe.address().port;
    await new Promise((ok) => probe.close(ok));
    const config = join(root, 'runtime fixture.json');
    await writeFile(
      config,
      JSON.stringify({ siteDirectory: join(root, 'site'), httpPort: port, office: false }),
    );
    const systemPath =
      process.platform === 'win32' ? join(process.env.WINDIR, 'System32') : '/nonexistent';
    child = spawn(process.execPath, [join(root, 'app/server.mjs'), config], {
      cwd: root,
      env: { ...process.env, PATH: systemPath, NODE_PATH: '', NODE_OPTIONS: '' },
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
    });
    const closed = once(child, 'close');
    let stderr = '';
    child.stderr.setEncoding('utf8').on('data', (chunk) => {
      stderr += chunk;
    });
    const ready = await new Promise((ok, fail) => {
      let stdout = '';
      child.once('error', fail);
      child.once('exit', (code) =>
        fail(new Error(`Portable child exited early (${code}): ${stderr}`)),
      );
      child.stdout.setEncoding('utf8').on('data', (chunk) => {
        stdout += chunk;
        if (stdout.includes('\n')) {
          try {
            ok(JSON.parse(stdout.split('\n')[0]));
          } catch (error) {
            fail(error);
          }
        }
      });
    });
    assert.equal(ready.ready, true);
    assert.equal(ready.httpPort, port);
    assert.equal(
      (await (await fetch(`http://127.0.0.1:${port}/health`)).json()).product,
      'office-helper-portable',
    );
    assert.equal(
      await (await fetch(`http://127.0.0.1:${port}/`)).text(),
      '<h1>Isolated helper</h1>',
    );
    assert.equal((await fetch(`http://127.0.0.1:${port}/native-check.html`)).status, 404);
    child.stdin.end('stop\n');
    assert.deepEqual(await closed, [0, null]);
    assert.equal(stderr, '');
    const released = http.createServer();
    await new Promise((ok, fail) => {
      released.once('error', fail);
      released.listen(port, '127.0.0.1', ok);
    });
    await new Promise((ok) => released.close(ok));
  },
);
