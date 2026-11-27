// Opt-in verification of an already-generated archive. No launcher/controller calls.
import assert from 'node:assert/strict';
import { createHash, X509Certificate } from 'node:crypto';
import { spawn, execFile } from 'node:child_process';
import http from 'node:http';
import https from 'node:https';
import net from 'node:net';
import { readFile, writeFile, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve, delimiter } from 'node:path';
import { promisify } from 'node:util';

const artifact = JSON.parse(
  await readFile('output/verification/portable-022-artifact.json', 'utf8'),
);
const root = resolve(artifact.extractedRoot);
assert.match(root, /中文 空格/);
const executable = join(root, 'runtime/node.exe');
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
assert.equal(digest(await readFile(executable)), digest(await readFile(process.execPath)));
const env = { ...process.env };
for (const key of Object.keys(env)) {
  if (/^(PATH|NODE_.*|OPENSSL_CONF|SSL_CERT_FILE|SSL_CERT_DIR)$/i.test(key)) delete env[key];
}
env.PATH = [join(process.env.WINDIR, 'System32'), process.env.WINDIR].join(delimiter);
env.NODE_OPTIONS = '';
env.NODE_PATH = '';
const nodeVersion = (
  await promisify(execFile)(executable, ['--version'], { cwd: root, env, windowsHide: true })
).stdout.trim();
assert.equal(nodeVersion, artifact.buildInfo.node);

// Read only pre-existing material; never import certificate/controller code.
const state = join(process.env.LOCALAPPDATA, 'OfficeHelperPortable');
const certificates = join(state, 'certificates');
const certificateManifestPath = join(certificates, 'manifest.json');
const certificateManifest = JSON.parse(await readFile(certificateManifestPath, 'utf8'));
assert.match(certificateManifest.activeId, /^[a-f0-9]{32}$/);
const certificatePath = join(certificates, certificateManifest.activeId, 'localhost.crt');
const privateKeyPath = join(certificates, certificateManifest.activeId, 'localhost.key');
const protectedPaths = [certificateManifestPath, certificatePath, privateKeyPath];
const protectedHashes = await Promise.all(
  protectedPaths.map(async (p) => digest(await readFile(p))),
);
const certificateBytes = await readFile(certificatePath);
const certificate = new X509Certificate(certificateBytes);
assert.ok(Date.parse(certificate.validFrom) <= Date.now());
assert.ok(Date.parse(certificate.validTo) > Date.now());
assert.equal(certificate.checkHost('localhost'), 'localhost');
assert.equal(certificate.checkIP('127.0.0.1'), '127.0.0.1');

const bindings = [
  { host: '127.0.0.1', port: 18765 },
  { host: '127.0.0.1', port: 18766 },
  { host: '::1', port: 18766 },
];
async function proveAvailable({ host, port }) {
  const server = net.createServer();
  await new Promise((ok, fail) => {
    server.once('error', fail);
    server.listen({ host, port, exclusive: true, ipv6Only: host === '::1' }, ok);
  });
  await new Promise((ok, fail) => server.close((error) => (error ? fail(error) : ok())));
}
for (const binding of bindings) await proveAvailable(binding);

const configDirectory = await mkdtemp(join(tmpdir(), 'OfficeHelper-022-运行 配置-'));
const config = join(configDirectory, 'runtime.json');
await writeFile(
  config,
  JSON.stringify({
    siteDirectory: join(root, 'site'),
    httpPort: 18765,
    httpsPort: 18766,
    office: true,
    certificatePath,
    privateKeyPath,
  }),
  { flag: 'wx' },
);

const checks = [];
function request(protocol, hostname, path, expectedStatus, expectedBytes) {
  const encrypted = protocol === 'https';
  return new Promise((ok, fail) => {
    const req = (encrypted ? https : http).request(
      {
        hostname,
        port: encrypted ? 18766 : 18765,
        path,
        agent: false,
        ...(encrypted
          ? {
              ca: certificateBytes,
              rejectUnauthorized: true,
              // The existing cert covers localhost and IPv4; IPv6 connects to
              // loopback while verifying the certificate identity localhost.
              ...(hostname === '::1' ? { servername: 'localhost' } : {}),
            }
          : {}),
      },
      (response) => {
        const chunks = [];
        response.on('data', (chunk) => chunks.push(chunk));
        response.once('error', fail);
        response.once('end', () => {
          try {
            const body = Buffer.concat(chunks);
            assert.equal(response.statusCode, expectedStatus, `${protocol} ${hostname} ${path}`);
            if (expectedBytes) assert.equal(digest(body), digest(expectedBytes));
            if (path === '/health') {
              assert.equal(JSON.parse(body).product, 'office-helper-portable');
            }
            checks.push({
              protocol,
              hostname,
              path,
              status: response.statusCode,
              ...(encrypted
                ? {
                    certificateVerified: response.socket.authorized,
                    verifiedIdentity: hostname === '::1' ? 'localhost' : hostname,
                  }
                : {}),
            });
            if (encrypted) assert.equal(response.socket.authorized, true);
            ok();
          } catch (error) {
            fail(error);
          }
        });
      },
    );
    req.setTimeout(4000, () => req.destroy(new Error('Packaged service request timed out')));
    req.once('error', fail);
    req.end();
  });
}

const withDeadline = async (promise, milliseconds, label) => {
  let timer;
  try {
    return await Promise.race([
      promise,
      new Promise((_, fail) => {
        timer = setTimeout(() => fail(new Error(label)), milliseconds);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
};
const child = spawn(executable, [join(root, 'app/server.mjs'), config], {
  cwd: root,
  env,
  stdio: ['pipe', 'pipe', 'pipe'],
  windowsHide: true,
});
let stdout = '';
let stderr = '';
child.stdout.setEncoding('utf8');
child.stderr.setEncoding('utf8').on('data', (chunk) => (stderr += chunk));
const closed = new Promise((ok) => child.once('close', (code, signal) => ok({ code, signal })));
let stopped = false;
try {
  const ready = await withDeadline(
    new Promise((ok, fail) => {
      child.once('error', fail);
      child.once('exit', (code) => fail(new Error(`Packaged server exited early: ${code}`)));
      child.stdout.on('data', (chunk) => {
        stdout += chunk;
        if (!stdout.includes('\n')) return;
        try {
          ok(JSON.parse(stdout.split('\n')[0]));
        } catch (error) {
          fail(error);
        }
      });
    }),
    10000,
    'Packaged server readiness timed out',
  );
  assert.deepEqual(ready, { ready: true, httpPort: 18765, httpsPort: 18766 });
  const publicPages = [
    ['/', 'index.html'],
    ['/?host=office', 'index.html'],
    ['/wps/writer/', 'wps/writer/index.html'],
    ['/wps/sheet/', 'wps/sheet/index.html'],
  ];
  const asset = artifact.buildInfo.siteFiles.find((f) => /^assets\/.+\.js$/.test(f.path)).path;
  publicPages.push(['/' + asset, asset]);
  const forbidden = [
    '/native-check.html',
    '/__native-report',
    '/extended-check.html',
    '/__extended-report',
    '/.local/native-token',
    '/src/main.tsx',
    '/wps/writer/native-check.html',
    '/wps/sheet/native-check.html',
    '/assets/app.js.map',
  ];
  for (const [protocol, hostname] of [
    ['http', '127.0.0.1'],
    ['https', 'localhost'],
  ]) {
    await request(protocol, hostname, '/health', 200);
    for (const [url, file] of publicPages) {
      await request(protocol, hostname, url, 200, await readFile(join(root, 'site', file)));
    }
    for (const url of forbidden) await request(protocol, hostname, url, 404);
  }
  for (const hostname of ['127.0.0.1', '::1']) await request('https', hostname, '/health', 200);
  child.stdin.end('stop\n');
  assert.deepEqual(await withDeadline(closed, 8000, 'Packaged service did not stop'), {
    code: 0,
    signal: null,
  });
  stopped = true;
  assert.equal(stderr, '');
  for (const binding of bindings) await proveAvailable(binding);
  const after = await Promise.all(protectedPaths.map(async (p) => digest(await readFile(p))));
  assert.deepEqual(after, protectedHashes, 'Existing certificate material changed');
  const result = {
    checkedAt: new Date().toISOString(),
    packageRoot: root,
    executable,
    nodeVersion,
    usedBundledNode: true,
    runtimeMatchesPackagingNode: true,
    pathContainsChineseAndSpaces: true,
    systemNodeAndNpmExcludedFromPath: true,
    systemPath: env.PATH,
    started: true,
    healthy: true,
    stopped: true,
    exitCode: 0,
    stderrEmpty: true,
    portsReleased: bindings,
    existingCertificateReusedReadOnly: true,
    certificateMaterialUnchanged: true,
    registrationPerformed: false,
    trustChangesPerformed: false,
    realOfficeOrWpsLaunched: false,
    checks,
    scope: 'Packaged service only. Candidate trial; no real host UI or new-feature acceptance.',
  };
  await writeFile(
    'output/verification/portable-022-runtime.json',
    JSON.stringify(result, null, 2),
    {
      flag: 'wx',
    },
  );
  console.log(
    JSON.stringify({ ...result, checks: `${checks.length} passed; see evidence file` }, null, 2),
  );
} finally {
  if (!stopped && child.exitCode === null && child.signalCode === null) {
    child.kill(); // Only this verifier's own child; never an existing port owner.
    await withDeadline(closed, 5000, 'Verifier child cleanup failed');
  }
}
