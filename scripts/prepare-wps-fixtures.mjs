import { mkdir, readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

// Dedicated, reproducible acceptance files; never overwrite a user's working documents.
const targetDirectory = 'output/native/wps-paragraphs';
await mkdir(targetDirectory, { recursive: true });
for (const extension of ['docx', 'xlsx']) {
  const zip = await JSZip.loadAsync(
    await readFile(`output/native/OfficeHelper-Native-Test.${extension}`),
  );
  if (extension === 'docx') {
    const part = 'word/document.xml';
    const xml = await zip.file(part).async('string');
    const paragraphs = ['OFFICE_HELPER_NATIVE_TEST', '第二段：居中验收', '第三段：保持原样']
      .map(
        (text, index) =>
          `<w:p><w:pPr><w:jc w:val="${index === 1 ? 'right' : 'left'}"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:sz w:val="36"/><w:b w:val="0"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`,
      )
      .join('');
    zip.file(
      part,
      xml.replace(
        /(<w:body[^>]*>)[\s\S]*?(<w:sectPr[\s\S]*?<\/w:sectPr>\s*)?<\/w:body>/,
        (_match, open, section) => open + paragraphs + (section ?? '') + '</w:body>',
      ),
    );
  } else {
    const part = Object.keys(zip.files).find((path) => /^xl\/worksheets\/[^/]+\.xml$/.test(path));
    const xml = await zip.file(part).async('string');
    const rows =
      '<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>OFFICE_HELPER_NATIVE_TEST</x:t></x:is></x:c></x:row>' +
      [120, 85, 45.5, 200]
        .map(
          (value, i) => `<x:row r="${i + 2}"><x:c r="B${i + 2}"><x:v>${value}</x:v></x:c></x:row>`,
        )
        .join('');
    zip.file(
      part,
      xml.replace(
        /<x:sheetData\s*\/>|<x:sheetData\b[^>]*>[\s\S]*?<\/x:sheetData>/,
        `<x:sheetData>${rows}</x:sheetData>`,
      ),
    );
  }
  const target = `${targetDirectory}/OfficeHelper-Native-Test.${extension}`;
  await writeFile(target, await zip.generateAsync({ type: 'nodebuffer' }));
  console.log(target);
}
