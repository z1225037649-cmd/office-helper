import { mkdir, readFile, writeFile, copyFile } from 'node:fs/promises';
import { resolve, join, dirname } from 'node:path';
import { promisify } from 'node:util';
import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { generateSideloadFile } from 'office-addin-dev-settings';
import { OfficeAddinManifest, OfficeApp } from 'office-addin-manifest';
import JSZip from 'jszip';
import {
  assertPortableOutputsAbsent,
  assertSafePortableContent,
  copyInspectedPortableSite,
  createPortableManifest,
  inspectPortableSite,
  populateWorksheetSample,
  portablePackagePaths,
} from '../portable/package-policy.mjs';

const run = promisify(execFile);
if (process.platform !== 'win32' || process.arch !== 'x64')
  throw new Error('Build the Windows x64 portable package on Windows x64.');
const pkg = JSON.parse(await readFile('package.json', 'utf8'));
const output = resolve('output/portable');
const paths = portablePackagePaths(output, pkg.version);
const { name, destination, archive } = paths;
// Check before writing anything. Never remove or replace an earlier deliverable.
await assertPortableOutputsAbsent(paths);
const siteFiles = await inspectPortableSite('dist');
const original = await readFile('manifests/office/manifest.xml', 'utf8');
const portableManifest = createPortableManifest(original, pkg.version);
for (const file of [
  'server.mjs',
  'site-policy.mjs',
  'controller.ps1',
  'certificates.ps1',
  'Launcher.cs',
  'launcher.manifest',
  'Launcher.exe.config',
  '使用说明.txt',
]) {
  assertSafePortableContent(await readFile(join('portable', file)), `portable/${file}`);
}
await mkdir(output, { recursive: true });
await mkdir(destination);
for (const part of ['app', 'runtime', 'site', 'manifests', 'samples', 'licenses'])
  await mkdir(join(destination, part), { recursive: true });
await copyInspectedPortableSite('dist', join(destination, 'site'), siteFiles);
for (const file of ['server.mjs', 'site-policy.mjs', 'controller.ps1', 'certificates.ps1']) {
  const bytes = await readFile(join('portable', file));
  assertSafePortableContent(bytes, `app/${file}`);
  await writeFile(join(destination, 'app', file), bytes, { flag: 'wx' });
}
await copyFile(process.execPath, join(destination, 'runtime/node.exe'));
const nodeLicense = join(dirname(process.execPath), 'LICENSE');
try {
  await copyFile(nodeLicense, join(destination, 'licenses/Node-LICENSE.txt'));
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
  const license = await fetch(
    `https://raw.githubusercontent.com/nodejs/node/${process.version}/LICENSE`,
  );
  if (!license.ok) throw new Error('Cannot obtain matching Node.js license: ' + license.status);
  await writeFile(join(destination, 'licenses/Node-LICENSE.txt'), await license.text());
}
for (const dependency of [
  'react',
  'react-dom',
  'scheduler',
  'lucide-react',
  'office-addin-dev-settings',
]) {
  try {
    await copyFile(
      join('node_modules', dependency, 'LICENSE'),
      join(destination, 'licenses', dependency + '-LICENSE.txt'),
    );
  } catch (error) {
    if (dependency !== 'office-addin-dev-settings' || error.code !== 'ENOENT') throw error;
    const license = await fetch(
      'https://raw.githubusercontent.com/OfficeDev/Office-Addin-Scripts/master/LICENSE',
    );
    if (!license.ok) throw new Error('Cannot obtain Office document template license');
    await writeFile(
      join(destination, 'licenses', dependency + '-LICENSE.txt'),
      await license.text(),
    );
  }
}
// Office requires manifest versions >= 1.0, even for a 0.x product release.
const manifestPath = join(destination, 'manifests/office.xml');
await writeFile(manifestPath, portableManifest);
for (const [component, hostName] of [
  ['writer', 'OfficeHelperPortableWriter'],
  ['sheet', 'OfficeHelperPortableSheet'],
]) {
  const folder = join(destination, 'site/wps', component);
  const ribbon = await readFile(join(folder, 'ribbon.xml'), 'utf8');
  await writeFile(
    join(folder, 'ribbon.xml'),
    ribbon
      .replaceAll('OfficeHelper', 'OfficeHelperPortable')
      .replace('label="小帮手"', 'label="小帮手（便携版）"'),
  );
  await writeFile(
    join(folder, 'manifest.xml'),
    `<?xml version="1.0" encoding="UTF-8"?><JsPlugin><ApiVersion>1.0.0</ApiVersion><Name>${hostName}</Name><Description>办公小帮手便携版</Description></JsPlugin>`,
  );
}
const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
for (const [app, extension] of [
  [OfficeApp.Word, 'docx'],
  [OfficeApp.Excel, 'xlsx'],
]) {
  const generated = await generateSideloadFile(app, manifest);
  const zip = await JSZip.loadAsync(await readFile(generated));
  if (extension === 'docx') {
    const part = 'word/document.xml';
    const document = zip.file(part);
    if (!document) throw new Error('The Microsoft Word template has no document.');
    const xml = await document.async('string');
    const paragraphFormat =
      '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr>';
    const body =
      `<w:p>${paragraphFormat}<w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:t>社区活动通知</w:t></w:r></w:p>` +
      `<w:p>${paragraphFormat}<w:r><w:t>请选中上方标题，在右侧小帮手预览并确认字号、加粗、段落居中、行距或首行缩进，再点击“恢复上一次修改”。例如：行距设为1.5倍、首行缩进1厘米。</w:t></w:r></w:p>`;
    const replaced = xml.replace(
      /(<w:body[^>]*>)[\s\S]*?(<w:sectPr[\s\S]*?<\/w:sectPr>\s*)?<\/w:body>/,
      (_match, open, section) => open + body + (section ?? '') + '</w:body>',
    );
    if (replaced === xml || !replaced.includes(body))
      throw new Error('The Microsoft Word template body could not be populated.');
    zip.file(part, replaced);
  } else {
    const part = Object.keys(zip.files).find((path) => /^xl\/worksheets\/[^/]+\.xml$/.test(path));
    if (!part) throw new Error('The Microsoft Excel template has no worksheet.');
    const xml = await zip.file(part).async('string');
    zip.file(part, populateWorksheetSample(xml));
  }
  await writeFile(
    join(destination, 'samples', 'OfficeHelper-Sample.' + extension),
    await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
  );
}
const compiler = join(process.env.WINDIR, 'Microsoft.NET/Framework64/v4.0.30319/csc.exe');
const iconImage = await readFile(join(destination, 'site/assets/icon-64.png'));
const iconHeader = Buffer.alloc(22);
iconHeader.writeUInt16LE(1, 2);
iconHeader.writeUInt16LE(1, 4);
iconHeader[6] = 64;
iconHeader[7] = 64;
iconHeader.writeUInt16LE(1, 10);
iconHeader.writeUInt16LE(32, 12);
iconHeader.writeUInt32LE(iconImage.length, 14);
iconHeader.writeUInt32LE(22, 18);
const iconPath = join(destination, 'app/helper.ico');
await writeFile(iconPath, Buffer.concat([iconHeader, iconImage]));
await run(compiler, [
  '/nologo',
  '/target:winexe',
  '/platform:anycpu',
  '/optimize+',
  '/utf8output',
  '/reference:System.Windows.Forms.dll',
  '/reference:System.Drawing.dll',
  '/reference:System.Web.Extensions.dll',
  '/win32manifest:' + resolve('portable/launcher.manifest'),
  '/win32icon:' + iconPath,
  '/out:' + join(destination, '办公小帮手.exe'),
  resolve('portable/Launcher.cs'),
]);
await copyFile('portable/Launcher.exe.config', join(destination, '办公小帮手.exe.config'));
await copyFile('portable/使用说明.txt', join(destination, '使用说明.txt'));
await writeFile(
  join(destination, 'build-info.json'),
  JSON.stringify(
    {
      version: pkg.version,
      node: process.version,
      platform: 'win32-x64',
      builtAt: new Date().toISOString(),
      siteFiles: await inspectPortableSite(join(destination, 'site')),
    },
    null,
    2,
  ),
);
await writeFile(
  join(destination, 'licenses/THIRD-PARTY.txt'),
  'Runtime: Node.js ' +
    process.version +
    '\nUI bundles: React, React DOM, scheduler, lucide-react.\nSample document templates: Microsoft Office-Addin-Scripts (MIT).\nSee accompanying license files. The portable server and launcher do not require npm.\n',
);
const powershell = join(process.env.WINDIR, 'System32/WindowsPowerShell/v1.0/powershell.exe');
// The legacy ZIP helper replaces its destination; fail closed before invoking it.
await assertPortableOutputsAbsent(paths, { includeDirectory: false });
// Paths are structured arguments, never interpolated into a shell expression.
await run(
  powershell,
  [
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    resolve('scripts/zip-portable.ps1'),
    '-Source',
    destination,
    '-Destination',
    archive,
  ],
  { maxBuffer: 1024 * 1024 },
);
const hash = createHash('sha256')
  .update(await readFile(archive))
  .digest('hex');
await writeFile(paths.checksum, hash + '  ' + name + '.zip\n', { flag: 'wx' });
console.log(JSON.stringify({ folder: destination, zip: archive, sha256: hash }, null, 2));
