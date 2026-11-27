import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import {
  assertPortableOutputsAbsent,
  assertSafePortableContent,
  copyInspectedPortableSite,
  createPortableManifest,
  inspectPortableSite,
  populateWorksheetSample,
  portablePackagePaths,
} from '../../portable/package-policy.mjs';
import { isPortableSitePath, requiredPortableSiteFiles } from '../../portable/site-policy.mjs';

async function fixture(t) {
  const root = await mkdtemp(join(tmpdir(), 'office-helper-package-中文 空格-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  return root;
}

async function write(root, path, content = 'fixture') {
  const filename = join(root, path);
  await mkdir(dirname(filename), { recursive: true });
  await writeFile(filename, content);
}

async function makeSite(root) {
  for (const path of [...requiredPortableSiteFiles, 'assets/index-fixture.js', 'assets/index.css'])
    await write(root, path);
}

test('portable release paths map 0.2.2 to a separate destination', () => {
  const paths = portablePackagePaths('output/portable', '0.2.2');
  assert.equal(paths.name, 'OfficeHelper-0.2.2-win-x64');
  assert.equal(paths.destination, resolve('output/portable/OfficeHelper-0.2.2-win-x64'));
  assert.equal(paths.archive, paths.destination + '.zip');
  assert.equal(paths.checksum, paths.archive + '.sha256');
});

test('portable versions reject traversal, prereleases and invalid Office components', () => {
  for (const version of [
    null,
    '',
    '../0.2.1',
    '0.2.2/../../old',
    '0.2.2-beta',
    '0.2',
    '01.2.2',
    '0.2.65536',
  ])
    assert.throws(() => portablePackagePaths('.', version), /version/);
});

test('each existing output blocks a release and is preserved byte for byte', async (t) => {
  const root = await fixture(t);
  for (const key of ['destination', 'archive', 'checksum']) {
    const output = join(root, key);
    await mkdir(output);
    const paths = portablePackagePaths(output, '0.2.2');
    const sentinel = key === 'destination' ? join(paths.destination, 'keep.txt') : paths[key];
    if (key === 'destination') await mkdir(paths.destination);
    await writeFile(sentinel, `existing-${key}`);
    await assert.rejects(assertPortableOutputsAbsent(paths), /already exists/);
    assert.equal(await readFile(sentinel, 'utf8'), `existing-${key}`);
  }
});

test('an old 0.2.1 release does not block 0.2.2 or change its bytes', async (t) => {
  const root = await fixture(t);
  const old = portablePackagePaths(root, '0.2.1');
  await mkdir(old.destination);
  await writeFile(old.archive, 'old-zip-sentinel');
  await writeFile(old.checksum, 'old-checksum-sentinel');
  await assertPortableOutputsAbsent(portablePackagePaths(root, '0.2.2'));
  assert.equal(await readFile(old.archive, 'utf8'), 'old-zip-sentinel');
  assert.equal(await readFile(old.checksum, 'utf8'), 'old-checksum-sentinel');
});

test('the public file policy excludes diagnostics, secrets, source maps and unknown WPS entries', () => {
  for (const path of requiredPortableSiteFiles) assert.ok(isPortableSitePath(path), path);
  for (const path of [
    'native-check.html',
    'src/main.tsx',
    '.local/native-token',
    'assets/app.js.map',
    'assets/private.key',
    'assets/native-check.js',
    'assets/debug.js',
    'assets/token.js',
    'assets/private.png',
    'assets/app.test.js',
    'assets/.env',
    'assets/../secret.js',
    'assets/subdir/app.js',
    'assets\\outside.js',
    'wps/writer/native-check.html',
    'wps/writer/token.js',
    'wps/localhost.pfx',
    'wps/extra.xml',
  ])
    assert.equal(isPortableSitePath(path), false, path);
});

test('clean built files are inspected, hashed and copied to an isolated site', async (t) => {
  const root = await fixture(t);
  const source = join(root, 'source');
  const target = join(root, 'isolated site');
  await makeSite(source);
  const inspected = await inspectPortableSite(source);
  assert.ok(inspected.every((file) => /^[a-f0-9]{64}$/.test(file.sha256)));
  await copyInspectedPortableSite(source, target, inspected);
  assert.deepEqual(await inspectPortableSite(target), inspected);
});

test('site inspection rejects an incomplete build', async (t) => {
  const root = await fixture(t);
  await write(root, 'index.html');
  await assert.rejects(inspectPortableSite(root), /Missing portable site entry/);
});

test('site inspection requires a compiled JavaScript bundle', async (t) => {
  const root = await fixture(t);
  for (const path of requiredPortableSiteFiles) await write(root, path);
  await assert.rejects(inspectPortableSite(root), /no compiled JavaScript bundle/);
});

for (const path of [
  'native-check.html',
  'assets/index.js.map',
  'assets/localhost.key',
  '.local/token',
]) {
  test(`site inspection fails closed for forbidden file ${path}`, async (t) => {
    const root = await fixture(t);
    await makeSite(root);
    await write(root, path);
    await assert.rejects(inspectPortableSite(root), /Unexpected portable site/);
  });
}

const forbiddenContents = [
  'fetch("/__native-report")',
  'open("/native-check.html")',
  'X-Helper-Test',
  'OFFICE_HELPER_NATIVE_TEST',
  '.local/native-token',
  '//# sourceMappingURL=app.js.map',
  '//# sourceMappingURL=data:application/json;base64,e30=',
  '?token=hidden',
  '-----BEGIN PRIVATE KEY-----',
  '-----BEGIN RSA PRIVATE KEY-----',
  '-----BEGIN EC PRIVATE KEY-----',
  '-----BEGIN OPENSSH PRIVATE KEY-----',
  'a'.repeat(48),
];

test('the content gate rejects diagnostic channels, tokens and PEMs without printing secrets', () => {
  for (const content of forbiddenContents) {
    for (const encoding of ['utf8', 'utf16le']) {
      assert.throws(
        () => assertSafePortableContent(Buffer.from(content, encoding), 'assets/app.js'),
        (error) => error.message.endsWith('assets/app.js') && !error.message.includes(content),
      );
    }
  }
  // Undo receipts are legitimate runtime tokens; SHA256 integrity metadata is public.
  assert.doesNotThrow(() =>
    assertSafePortableContent(Buffer.from('receipt.token; ' + 'b'.repeat(64)), 'app.js'),
  );
});

test('a safe filename cannot hide a diagnostic token inside a bundle', async (t) => {
  const root = await fixture(t);
  await makeSite(root);
  await write(root, 'assets/index-fixture.js', `const hidden="${'a'.repeat(48)}";`);
  await assert.rejects(inspectPortableSite(root), /forbidden/);
});

test('site inspection rejects directory junctions without following them', async (t) => {
  const root = await fixture(t);
  const site = join(root, 'site');
  const outside = join(root, 'outside');
  await mkdir(site);
  await mkdir(outside);
  await symlink(outside, join(site, 'assets'), process.platform === 'win32' ? 'junction' : 'dir');
  await assert.rejects(inspectPortableSite(site), /cannot contain links/);
});

test('site files changing between inspection and copying abort the copy', async (t) => {
  const root = await fixture(t);
  const source = join(root, 'source');
  await makeSite(source);
  const files = await inspectPortableSite(source);
  await write(source, 'assets/index-fixture.js', 'changed after inspection');
  await assert.rejects(
    copyInspectedPortableSite(source, join(root, 'target'), files),
    /changed after preflight/,
  );
});

test('copying an inspected site never overwrites existing destination files', async (t) => {
  const root = await fixture(t);
  const source = join(root, 'source');
  const target = join(root, 'target');
  await makeSite(source);
  await makeSite(target);
  const files = await inspectPortableSite(source);
  await assert.rejects(copyInspectedPortableSite(source, target, files), { code: 'EEXIST' });
});

test('portable manifest uses 1.0.2.2, the independent ID and the portable port', async () => {
  const original = await readFile(
    new URL('../../manifests/office/manifest.xml', import.meta.url),
    'utf8',
  );
  const manifest = createPortableManifest(original, '0.2.2');
  assert.match(manifest, /<Version>1\.0\.2\.2<\/Version>/);
  assert.match(manifest, /<Id>f5202c69-1f59-4e09-b510-829a3748a31b<\/Id>/);
  assert.match(manifest, /localhost:18766/);
  assert.doesNotMatch(manifest, /localhost:3443|开发版/);
  assert.match(manifest, /平均值.*重复漏填.*打印.*PDF.*通知/);
  assert.match(original, /<Version>1\.0\.0\.0<\/Version>/);
});

test('manifest preparation rejects missing or duplicate version and ID targets', () => {
  for (const original of [
    '<OfficeApp/>',
    '<Version>1.0.0.0</Version>',
    '<Id>26b5f5ec-853a-4c06-9c1c-8ac9ea33421a</Id><Version>1</Version><Version>2</Version>',
  ])
    assert.throws(() => createPortableManifest(original, '0.2.2'), /exactly one/);
});

for (const [label, xml] of [
  ['self-closing', '<worksheet><sheetData/></worksheet>'],
  ['empty', '<worksheet><sheetData></sheetData></worksheet>'],
  ['populated', '<worksheet><sheetData><row r="99"/></sheetData></worksheet>'],
  ['prefixed', '<x:worksheet xmlns:x="urn:sheet"><x:sheetData/></x:worksheet>'],
  ['commented decoy', '<worksheet><!--<sheetData/>--><sheetData/></worksheet>'],
]) {
  test(`portable Excel sample populates ten addressed cells: ${label}`, () => {
    const result = populateWorksheetSample(xml);
    assert.equal((result.match(/<(?:x:)?row\b/g) ?? []).length, 5);
    assert.equal((result.match(/<(?:x:)?c\b/g) ?? []).length, 10);
    for (const address of ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4', 'A5', 'B5'])
      assert.ok(result.includes(`r="${address}"`), address);
    const values = [...result.matchAll(/<(?:x:)?v>([\d.]+)<\/(?:x:)?v>/g)].map((match) =>
      Number(match[1]),
    );
    assert.deepEqual(values, [120, 85, 45.5, 200]);
    assert.equal(
      values.reduce((total, value) => total + value, 0),
      450.5,
    );
    assert.equal(values.reduce((total, value) => total + value, 0) / values.length, 112.625);
    if (label === 'commented decoy') assert.ok(result.includes('<!--<sheetData/>-->'));
  });
}

test('sample dimension updates preserve attributes and ignore comment decoys', () => {
  const result = populateWorksheetSample(
    '<worksheet><!--<dimension ref="Z99"/>--><dimension other="keep" ref=\'A1\'></dimension><sheetData/></worksheet>',
  );
  assert.match(result, /<!--<dimension ref="Z99"\/>-->/);
  assert.match(result, /<dimension other="keep" ref="A1:B5"\/>/);
});

test('missing or ambiguous sample data is rejected', () => {
  for (const xml of [
    '<worksheet/>',
    '<worksheet><!--<sheetData/>--></worksheet>',
    '<worksheet><sheetData/><sheetData/></worksheet>',
  ])
    assert.throws(() => populateWorksheetSample(xml), /exactly one/);
});

test('package and lockfile root versions agree on 0.2.2', async () => {
  const pkg = JSON.parse(await readFile(new URL('../../package.json', import.meta.url), 'utf8'));
  const lock = JSON.parse(
    await readFile(new URL('../../package-lock.json', import.meta.url), 'utf8'),
  );
  assert.equal(pkg.version, '0.2.2');
  assert.equal(lock.version, pkg.version);
  assert.equal(lock.packages[''].version, pkg.version);
});
