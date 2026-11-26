import { mkdir, readFile, writeFile, copyFile } from 'node:fs/promises';
import { randomBytes } from 'node:crypto';
import { resolve } from 'node:path';
import { generateSideloadFile } from 'office-addin-dev-settings';
import { OfficeAddinManifest, OfficeApp } from 'office-addin-manifest';

await mkdir('.local', { recursive: true });
await mkdir('output/native', { recursive: true });
const token = randomBytes(24).toString('hex');
await writeFile('.local/native-token', token);
const original = await readFile('manifests/office/manifest.xml', 'utf8');
const testManifest = original.replace(
  'https://localhost:3443/?host=office',
  `https://localhost:3443/native-check.html?host=office&amp;token=${token}`,
);
await writeFile('.local/office-native-check.xml', testManifest);
const manifest = await OfficeAddinManifest.readManifestFile(
  resolve('.local/office-native-check.xml'),
);
for (const [app, filename] of [
  [OfficeApp.Word, 'OfficeHelper-Native-Test.docx'],
  [OfficeApp.Excel, 'OfficeHelper-Native-Test.xlsx'],
]) {
  const temp = await generateSideloadFile(app, manifest);
  await copyFile(temp, `output/native/${filename}`);
}
await mkdir('.local/wps-native-test', { recursive: true });
await copyFile('public/wps/ribbon.xml', '.local/wps-native-test/ribbon.xml');
await writeFile(
  '.local/wps-native-test/index.html',
  '<!doctype html><html><head><meta charset="UTF-8"></head><body><script src="main.js"></script></body></html>',
);
await import('./prepare-wps-native.mjs');
console.log('已生成独立测试文件和本地测试入口。');
