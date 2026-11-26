import { mkdir, readFile, writeFile, copyFile } from 'node:fs/promises';
import { randomBytes } from 'node:crypto';
import { resolve } from 'node:path';
import { generateSideloadFile } from 'office-addin-dev-settings';
import { OfficeAddinManifest, OfficeApp } from 'office-addin-manifest';
import JSZip from 'jszip';

const wordFormatOnly = process.argv.includes('--word-format');
await mkdir('.local', { recursive: true });
await mkdir('output/native', { recursive: true });
// A focused fixture regeneration must not invalidate an active WPS test session.
const previousToken = wordFormatOnly
  ? await readFile('.local/native-token', 'utf8').catch(() => '')
  : '';
const token = /^[a-f0-9]{48}$/.test(previousToken.trim())
  ? previousToken.trim()
  : randomBytes(24).toString('hex');
if (token !== previousToken.trim()) await writeFile('.local/native-token', token);
const original = await readFile('manifests/office/manifest.xml', 'utf8');
const nativeUrl = `https://localhost:3443/native-check.html?host=office&token=${token}`;
function testManifest(url) {
  const source = original.replace(
    'https://localhost:3443/?host=office',
    url.replaceAll('&', '&amp;'),
  );
  if (source === original) throw new Error('没有找到Office开发入口，未生成验收清单');
  return source;
}

async function prepareWordFormat() {
  const manifestPath = '.local/office-word-format-check.xml';
  const url = `${nativeUrl}&check=paragraph-format`;
  await writeFile(manifestPath, testManifest(url));
  const manifest = await OfficeAddinManifest.readManifestFile(resolve(manifestPath));
  const temp = await generateSideloadFile(OfficeApp.Word, manifest);
  const zip = await JSZip.loadAsync(await readFile(temp));
  const part = zip.file('word/document.xml');
  if (!part) throw new Error('Word侧载模板缺少正文，未生成格式验收样本');
  const xml = await part.async('string');
  const originals = [
    {
      text: 'OFFICE_HELPER_NATIVE_TEST',
      alignment: 'left',
      line: 240,
      left: 120,
      right: 160,
      indent: '',
    },
    {
      text: '第二段：居中验收',
      alignment: 'right',
      line: 360,
      left: 0,
      right: 80,
      indent: ' w:firstLine="240"',
    },
    {
      text: '第三段：保持原样',
      alignment: 'left',
      line: 480,
      left: 240,
      right: 0,
      indent: ' w:hanging="240"',
    },
  ];
  const runProperties =
    '<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>';
  const paragraphs = originals
    .map(
      (p) =>
        `<w:p><w:pPr><w:spacing w:line="${p.line}" w:lineRule="auto"/><w:ind w:left="${p.left}" w:right="${p.right}"${p.indent}/><w:jc w:val="${p.alignment}"/>${runProperties}</w:pPr><w:r>${runProperties}<w:t>${p.text}</w:t></w:r></w:p>`,
    )
    .join('');
  const replaced = xml.replace(
    /(<w:body[^>]*>)[\s\S]*?(<w:sectPr[\s\S]*?<\/w:sectPr>\s*)?<\/w:body>/,
    (_match, open, section) => open + paragraphs + (section ?? '') + '</w:body>',
  );
  if (replaced === xml) throw new Error('Word侧载模板正文结构未知，未生成格式验收样本');
  zip.file('word/document.xml', replaced);
  const directory = 'output/native/word-format';
  await mkdir(directory, { recursive: true });
  const fixture = `${directory}/OfficeHelper-Native-Test.docx`;
  await writeFile(fixture, await zip.generateAsync({ type: 'nodebuffer' }));
  // Keep the token out of terminal output; this local file is the exact runnable URL.
  await writeFile('.local/word-format-check-url.txt', url + '\n');
  console.log(`Word专项样本：${fixture}`);
  console.log(`Word专项清单：${manifestPath}`);
  console.log('Word专项URL：.local/word-format-check-url.txt（须在真实Word任务窗格运行）');
}

if (!wordFormatOnly) {
  await writeFile('.local/native-token', token);
  await writeFile('.local/office-native-check.xml', testManifest(nativeUrl));
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
}
await prepareWordFormat();
console.log('已生成独立测试文件和本地测试入口；未启动宿主或执行实机验收。');
