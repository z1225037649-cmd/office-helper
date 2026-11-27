import type { SheetBatchSource, SheetCellValue } from './types';
import { validateSources } from './validate';

export type Delimiter = ',' | '\t';

function decodeText(bytes: ArrayBuffer): string {
  const data = new Uint8Array(bytes);
  if (data.length >= 3 && data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf)
    return new TextDecoder('utf-8', { fatal: true }).decode(data.subarray(3));
  if (data.length >= 2 && data[0] === 0xff && data[1] === 0xfe)
    return new TextDecoder('utf-16le', { fatal: true }).decode(data.subarray(2));
  if (data.length >= 2 && data[0] === 0xfe && data[1] === 0xff)
    return new TextDecoder('utf-16be', { fatal: true }).decode(data.subarray(2));

  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(data);
  } catch {
    try {
      return new TextDecoder('gb18030', { fatal: true }).decode(data);
    } catch {
      throw new TypeError('文件编码无法识别，请另存为 UTF-8 CSV 后重试。');
    }
  }
}

function countFirstRecord(text: string, delimiter: Delimiter): number {
  let quoted = false;
  let count = 0;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]!;
    if (character === '"') {
      if (quoted && text[index + 1] === '"') index += 1;
      else quoted = !quoted;
    } else if (!quoted && character === delimiter) {
      count += 1;
    } else if (!quoted && (character === '\r' || character === '\n')) {
      break;
    }
  }
  return count;
}

export function detectDelimiter(text: string, fileName = ''): Delimiter {
  if (/\.tsv$/iu.test(fileName)) return '\t';
  if (/\.csv$/iu.test(fileName)) return ',';
  const tabs = countFirstRecord(text, '\t');
  const commas = countFirstRecord(text, ',');
  return tabs > commas ? '\t' : ',';
}

function parseRows(text: string, delimiter: Delimiter): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  let closedQuote = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]!;
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          cell += '"';
          index += 1;
        } else {
          quoted = false;
          closedQuote = true;
        }
      } else {
        cell += character;
      }
      continue;
    }

    if (closedQuote) {
      if (character !== delimiter && character !== '\r' && character !== '\n')
        throw new TypeError('CSV 引号后包含多余字符，请检查文件格式。');
      closedQuote = false;
    }

    if (character === '"') {
      if (cell.length !== 0) throw new TypeError('CSV 字段中的引号位置不正确。');
      quoted = true;
    } else if (character === delimiter) {
      row.push(cell);
      cell = '';
    } else if (character === '\r' || character === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      if (character === '\r' && text[index + 1] === '\n') index += 1;
    } else {
      cell += character;
    }
  }

  if (quoted) throw new TypeError('CSV 引号没有闭合，请检查文件格式。');
  if (cell.length > 0 || row.length > 0 || closedQuote) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

export function parseDelimitedText(text: string, sourceLabel: string): SheetBatchSource {
  const normalized = text.replace(/^\uFEFF/u, '');
  if (!normalized.trim()) throw new TypeError(`文件“${sourceLabel}”没有内容。`);
  const delimiter = detectDelimiter(normalized, sourceLabel);
  const [headers, ...rawRows] = parseRows(normalized, delimiter);
  if (!headers?.length) throw new TypeError(`文件“${sourceLabel}”缺少表头。`);

  const rows: SheetCellValue[][] = [];
  for (let index = 0; index < rawRows.length; index += 1) {
    const row = rawRows[index]!;
    if (row.length > headers.length)
      throw new TypeError(`文件“${sourceLabel}”第 ${index + 2} 行比表头多出单元格。`);
    rows.push(row);
  }

  const source = { sourceLabel, headers, rows };
  validateSources([source]);
  return source;
}

export function parseDelimitedFile(bytes: ArrayBuffer, sourceLabel: string): SheetBatchSource {
  return parseDelimitedText(decodeText(bytes), sourceLabel);
}

function csvField(value: SheetCellValue): string {
  const text = value === null || value === undefined ? '' : String(value);
  return /[",\r\n]/u.test(text) ? `"${text.replace(/"/gu, '""')}"` : text;
}

/** Spreadsheet programs can evaluate formula-like text when opening CSV. Refuse unsafe export. */
export function formulaLikeTextCells(source: SheetBatchSource): number {
  let count = 0;
  const cells = [source.headers, ...source.rows];
  for (const row of cells)
    for (const value of row)
      if (typeof value === 'string') {
        const candidate = value.replace(/^[\s\u0000-\u0020]+/u, '');
        if (/^[=+@]/u.test(candidate) || /^-(?=\D|$)/u.test(candidate) || /^-\d+[=+*/^]/u.test(candidate))
          count += 1;
      }
  return count;
}

export function serializeSheetBatchCsv(
  columns: readonly string[],
  rows: readonly (readonly SheetCellValue[])[],
): string {
  return `\uFEFF${[columns, ...rows].map((row) => row.map(csvField).join(',')).join('\r\n')}\r\n`;
}
