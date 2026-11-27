import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import JSZip from 'jszip';

function maskXmlComments(xml) {
  return xml.replace(/<!--[\s\S]*?-->/g, (comment) => comment.replace(/[^\r\n]/g, ' '));
}

export function populateAverageWorksheet(xml) {
  const sheetDataPattern =
    /<((?:[A-Za-z_][\w.-]*:)?sheetData)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1\s*>)/g;
  const sheetDataMatches = [...maskXmlComments(xml).matchAll(sheetDataPattern)];
  if (sheetDataMatches.length !== 1)
    throw new Error('The Excel template must contain exactly one supported sheetData.');

  const sheetDataMatch = sheetDataMatches[0];
  const originalSheetData = xml.slice(
    sheetDataMatch.index,
    sheetDataMatch.index + sheetDataMatch[0].length,
  );
  const openingTag = originalSheetData.match(
    /^<((?:[A-Za-z_][\w.-]*:)?sheetData)\b([^>]*?)(?:\/>|>)/,
  );
  if (!openingTag) throw new Error('The Excel template has an unsupported sheetData node.');
  const tag = openingTag[1];
  const attributes = openingTag[2];
  const prefix = tag.slice(0, -'sheetData'.length);
  const cell = (address, value, type = '') =>
    `<${prefix}c r="${address}"${type}><${prefix}${type ? 'is' : 'v'}>${type ? `<${prefix}t>${value}</${prefix}t>` : value}</${prefix}${type ? 'is' : 'v'}></${prefix}c>`;
  const rows =
    `<${prefix}row r="1">${cell('A1', 'OFFICE_HELPER_NATIVE_TEST', ' t="inlineStr"')}${cell('B1', '金额', ' t="inlineStr"')}</${prefix}row>` +
    [120, 85, 45.5, 200]
      .map(
        (value, index) =>
          `<${prefix}row r="${index + 2}">${cell(`B${index + 2}`, value)}</${prefix}row>`,
      )
      .join('');
  let updated =
    xml.slice(0, sheetDataMatch.index) +
    `<${tag}${attributes}>${rows}</${tag}>` +
    xml.slice(sheetDataMatch.index + sheetDataMatch[0].length);

  const dimensionPattern = /<((?:[A-Za-z_][\w.-]*:)?dimension)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1\s*>)/;
  const maskedUpdated = maskXmlComments(updated);
  const dimensionMatch = dimensionPattern.exec(maskedUpdated);
  if (dimensionMatch) {
    const originalDimension = updated.slice(
      dimensionMatch.index,
      dimensionMatch.index + dimensionMatch[0].length,
    );
    const dimensionTag = originalDimension.match(/^<((?:[A-Za-z_][\w.-]*:)?dimension)\b/);
    const dimensionAttributes = originalDimension.match(/^<[^\s/>]+\b([^>]*)/);
    if (!dimensionTag || !dimensionAttributes)
      throw new Error('The Excel template has an unsupported dimension node.');
    const withoutRef = dimensionAttributes[1].replace(/\s+ref\s*=\s*(?:"[^"]*"|'[^']*')/i, '');
    const replacement = `<${dimensionTag[1]}${withoutRef} ref="A1:B5"/>`;
    updated =
      updated.slice(0, dimensionMatch.index) +
      replacement +
      updated.slice(dimensionMatch.index + dimensionMatch[0].length);
  } else {
    const rootPattern = /<((?:[A-Za-z_][\w.-]*:)?worksheet)\b[^>]*>/;
    const root = rootPattern.exec(maskedUpdated);
    if (root) {
      const prefix = root[1].match(/^([A-Za-z_][\w.-]*:)?/)?.[1] ?? '';
      const sheetProperties = new RegExp(
        `<${prefix}sheetPr\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${prefix}sheetPr\\s*>)`,
      );
      const properties = sheetProperties.exec(maskedUpdated);
      const insertionPoint = properties
        ? properties.index + properties[0].length
        : root.index + root[0].length;
      const dimension = `<${prefix}dimension ref="A1:B5"/>`;
      updated = updated.slice(0, insertionPoint) + dimension + updated.slice(insertionPoint);
    }
  }
  if (updated === xml) throw new Error('The Excel template must contain a worksheet root.');
  return updated;
}

export async function prepareWpsAverageFixture(
  source = 'output/native/OfficeHelper-Native-Test.xlsx',
  target = 'output/native/wps/OfficeHelper-Native-Test.xlsx',
) {
  const workbook = await JSZip.loadAsync(await readFile(source));
  const sheetPaths = Object.keys(workbook.files).filter((path) =>
    /^xl\/worksheets\/[^/]+\.xml$/.test(path),
  );
  if (sheetPaths.length !== 1)
    throw new Error('Expected exactly one worksheet in the generated test workbook.');
  const part = workbook.file(sheetPaths[0]);
  if (!part) throw new Error('The generated worksheet is missing.');
  workbook.file(sheetPaths[0], populateAverageWorksheet(await part.async('string')));
  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, await workbook.generateAsync({ type: 'nodebuffer' }));
  return target;
}

if (process.argv[1]?.endsWith('prepare-wps-average-fixture.mjs')) {
  console.log(await prepareWpsAverageFixture());
}
