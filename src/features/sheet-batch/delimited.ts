import type { SheetBatchSource, SheetCellValue } from './types';
import { validateSources } from './validate';

export type Delimiter = ',' | '\t';

/**
 * Bounds for a single delimited file. `maximumRows` counts data records but does not
 * count the header; `maximumCells` includes both the header and all data records.
 */
export interface DelimitedParseLimits {
  readonly maximumBytes: number;
  readonly maximumCharacters: number;
  readonly maximumRows: number;
  readonly maximumColumns: number;
  readonly maximumCells: number;
}

export type DelimitedParseOptions = Partial<DelimitedParseLimits>;

const defaultParseLimits: DelimitedParseLimits = {
  maximumBytes: 10 * 1024 * 1024,
  maximumCharacters: 10 * 1024 * 1024,
  maximumRows: 100_000,
  maximumColumns: 200,
  maximumCells: 1_000_000,
};

function resolveParseLimits(options: DelimitedParseOptions): DelimitedParseLimits {
  const limits = { ...defaultParseLimits, ...options };
  for (const [name, value] of Object.entries(limits)) {
    const minimum = name === 'maximumRows' ? 0 : 1;
    if (!Number.isSafeInteger(value) || value < minimum) {
      throw new TypeError(`CSV 解析限制 ${name} 必须是至少 ${minimum} 的安全整数。`);
    }
  }
  return limits;
}

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

function parseRows(
  text: string,
  delimiter: Delimiter,
  limits: DelimitedParseLimits,
  sourceLabel: string,
): { headers: string[] | undefined; rows: string[][] } {
  let headers: string[] | undefined;
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  let closedQuote = false;
  let parsedCellCount = 0;

  const pushCell = () => {
    if (headers !== undefined && row.length >= headers.length) {
      throw new TypeError(`文件“${sourceLabel}”第 ${rows.length + 2} 条记录比表头多出单元格。`);
    }
    if (row.length >= limits.maximumColumns) {
      throw new TypeError(`文件“${sourceLabel}”的列数超过 ${limits.maximumColumns} 列。`);
    }
    row.push(cell);
    parsedCellCount += 1;
    if (parsedCellCount > limits.maximumCells) {
      throw new TypeError(`文件“${sourceLabel}”的单元格总数超过 ${limits.maximumCells} 格。`);
    }
    cell = '';
  };

  const pushRow = () => {
    if (headers === undefined) {
      headers = row;
    } else if (rows.length >= limits.maximumRows) {
      throw new TypeError(`文件“${sourceLabel}”的数据记录数超过 ${limits.maximumRows} 条。`);
    } else {
      rows.push(row);
    }
    row = [];
  };

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
      pushCell();
    } else if (character === '\r' || character === '\n') {
      pushCell();
      pushRow();
      if (character === '\r' && text[index + 1] === '\n') index += 1;
    } else {
      cell += character;
    }
  }

  if (quoted) throw new TypeError('CSV 引号没有闭合，请检查文件格式。');
  if (cell.length > 0 || row.length > 0 || closedQuote) {
    pushCell();
    pushRow();
  }
  return { headers, rows };
}

export function parseDelimitedText(
  text: string,
  sourceLabel: string,
  options: DelimitedParseOptions = {},
): SheetBatchSource {
  const limits = resolveParseLimits(options);
  if (typeof text !== 'string') throw new TypeError('CSV 内容必须是文字。');
  if (text.length > limits.maximumCharacters) {
    throw new TypeError(`文件“${sourceLabel}”的文字长度超过 ${limits.maximumCharacters} 个字符。`);
  }
  const normalized = text.replace(/^\uFEFF/u, '');
  if (!normalized.trim()) throw new TypeError(`文件“${sourceLabel}”没有内容。`);
  const delimiter = detectDelimiter(normalized, sourceLabel);
  const { headers, rows } = parseRows(normalized, delimiter, limits, sourceLabel);
  if (!headers?.length) throw new TypeError(`文件“${sourceLabel}”缺少表头。`);

  const source: SheetBatchSource = { sourceLabel, headers, rows };
  validateSources([source]);
  return source;
}

export function parseDelimitedFile(
  bytes: ArrayBuffer,
  sourceLabel: string,
  options: DelimitedParseOptions = {},
): SheetBatchSource {
  const limits = resolveParseLimits(options);
  if (bytes.byteLength > limits.maximumBytes) {
    throw new TypeError(`文件“${sourceLabel}”超过 ${limits.maximumBytes} 字节。`);
  }
  return parseDelimitedText(decodeText(bytes), sourceLabel, limits);
}

function csvField(value: SheetCellValue): string {
  const text = value === null || value === undefined ? '' : String(value);
  return /[",\r\n\t]/u.test(text) ? `"${text.replace(/"/gu, '""')}"` : text;
}

function isFormulaLikeText(value: string): boolean {
  const candidate = value.replace(/^[\s\u0000-\u0020]+/u, '');
  if (/^[=+@]/u.test(candidate)) return true;
  if (!candidate.startsWith('-')) return false;

  // Preserve ordinary negative integers, decimals, and exponent-form numbers.
  if (/^-(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/iu.test(candidate)) return false;

  // Keep the conservative legacy handling for other leading minus text, since a
  // spaced unary operator can still precede a formula such as `- SUM(A1)`.
  return true;
}

/** Spreadsheet programs can evaluate formula-like text when opening CSV. Refuse unsafe export. */
export function formulaLikeTextCells(source: SheetBatchSource): number {
  let count = 0;
  const cells = [source.headers, ...source.rows];
  for (const row of cells)
    for (const value of row)
      if (typeof value === 'string') {
        if (isFormulaLikeText(value)) count += 1;
      }
  return count;
}

export function serializeSheetBatchCsv(
  columns: readonly string[],
  rows: readonly (readonly SheetCellValue[])[],
): string {
  return `\uFEFF${[columns, ...rows].map((row) => row.map(csvField).join(',')).join('\r\n')}\r\n`;
}
