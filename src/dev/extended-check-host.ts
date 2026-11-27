// Read-only observers. Feature writes, undo, and navigation belong to the real runner.
import { OfficeAdapter } from '../adapters/office/adapter';
import { WpsAdapter } from '../adapters/wps/adapter';
import { getWpsApplication, type WpsDocument, type WpsSelection } from '../adapters/wps/bridge';
import type { HostAdapter } from '../core/contracts';
import {
  assertSourcePath,
  extendedMarker,
  normalizeText,
  same,
  type ExtendedConfig,
} from './extended-check-spec';

export class ExtendedUnsupported extends Error {}
export interface ExtendedNativeState {
  path: string;
  text?: string;
  xml?: string;
  values?: unknown[][];
  formulas?: (string | null)[][];
  second?: unknown[][];
  catalogue?: string[];
  widths?: number[];
  bold?: boolean[][];
  wrap?: boolean[][];
  hidden?: boolean[];
}
export interface ExtendedObserver {
  read(): Promise<ExtendedNativeState>;
  notice(): Promise<{ source: ExtendedNativeState; text: string; independent: boolean }>;
}

function unsupported(message: string): never {
  throw new ExtendedUnsupported(message);
}
function bool(value: unknown): boolean {
  if (typeof value === 'boolean') return value;
  if (value === 0 || value === 1 || value === -1) return value !== 0;
  return unsupported('原生逐格格式属性缺失或混合，不能验证。');
}
function matrix(value: unknown, rows: number, columns: number): unknown[][] {
  if (
    !Array.isArray(value) ||
    value.length !== rows ||
    value.some((row) => !Array.isArray(row) || row.length !== columns)
  )
    return unsupported('原生矩阵维度不可靠。');
  return value.map((row) => [...row]);
}
const formula = (value: unknown): string | null =>
  typeof value === 'string' && value.startsWith('=') ? value.toUpperCase() : null;

export function assertNativeFixture(state: ExtendedNativeState, config: ExtendedConfig): void {
  assertSourcePath(state.path, config);
  if (config.kind === 'text') {
    if (state.text !== config.text) throw new Error('专用正文的完整标记不匹配。');
    return;
  }
  const values = state.values;
  const selected = values?.slice(2, 10).map((row) => row.slice(1, 6));
  const complete: unknown[][] = Array.from({ length: 12 }, () => Array<unknown>(6).fill(''));
  complete[0]!.splice(0, 3, extendedMarker, config.run, config.matrix);
  config.values.forEach((row, r) => complete[r + 2]!.splice(1, 5, ...row));
  complete[11]![1] = 'OFFICE_HELPER_EXTENDED_OUTSIDE';
  complete[11]![5] = 909;
  const formulas = Array.from({ length: 12 }, () => Array<string | null>(6).fill(null));
  formulas[7]![5] = '=SUM(F4:F5)';
  if (
    values?.length !== 12 ||
    values.some((row) => row.length !== 6) ||
    !same(values[0]?.slice(0, 3), [extendedMarker, config.run, config.matrix]) ||
    !same(selected, config.values) ||
    !same(
      values?.map((row) => row.map((v) => (v === null ? '' : v))),
      complete,
    ) ||
    !same(state.formulas, formulas) ||
    values[11]?.[1] !== 'OFFICE_HELPER_EXTENDED_OUTSIDE' ||
    values[11]?.[5] !== 909 ||
    !same(state.second, [['OFFICE_HELPER_EXTENDED_SECOND_SHEET', config.run, config.matrix]]) ||
    !same(state.catalogue, ['ExtendedData', 'ScopeSentinel']) ||
    state.hidden?.[8] !== true
  )
    throw new Error('完整原生矩阵、公式、第2表、隐藏行或样本标记不匹配。');
}

export function createExtendedObserver(
  host: HostAdapter,
  config: ExtendedConfig,
): ExtendedObserver {
  if (
    host.info.id !== config.host ||
    host.info.kind !== config.kind ||
    !(
      (host instanceof OfficeAdapter && config.host === 'office') ||
      (host instanceof WpsAdapter && config.host === 'wps')
    )
  )
    throw new ExtendedUnsupported('必须使用真实 Office/WPS 适配器；不接受 mock 或自定义适配器。');
  if (config.host === 'office') {
    if (
      typeof Office === 'undefined' ||
      Office.context.host !==
        (config.kind === 'text' ? Office.HostType.Word : Office.HostType.Excel)
    )
      throw new ExtendedUnsupported('不在指定真实 Office 任务窗格中。');
    const read = async (): Promise<ExtendedNativeState> => {
      const path = Office.context.document.url;
      assertSourcePath(path, config);
      if (config.kind === 'text')
        return Word.run(async (ctx) => {
          const body = ctx.document.body;
          body.load('text');
          const xml = body.getOoxml();
          await ctx.sync();
          return { path, text: normalizeText(body.text), xml: xml.value };
        });
      return Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load('items/name');
        await ctx.sync();
        if (
          !same(
            sheets.items.map((sheet) => sheet.name),
            ['ExtendedData', 'ScopeSentinel'],
          )
        )
          throw new Error('工作表目录不匹配。');
        const sheet = sheets.items[0]!;
        const second = sheets.items[1]!;
        const range = sheet.getRange('A1:F12');
        const other = second.getRange('A1:C1');
        const used = sheet.getUsedRange();
        const otherUsed = second.getUsedRange();
        used.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
        otherUsed.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
        range.load(['values', 'formulas']);
        other.load(['values', 'formulas']);
        const widths = Array.from({ length: 6 }, (_, c) => {
          const format = range.getColumn(c).format;
          format.load('columnWidth');
          return format;
        });
        const hidden = Array.from({ length: 12 }, (_, r) => {
          const row = range.getRow(r);
          row.load('rowHidden');
          return row;
        });
        const formats = Array.from({ length: 12 }, (_, r) =>
          Array.from({ length: 6 }, (_, c) => {
            const format = range.getCell(r, c).format;
            const font = format.font;
            format.load('wrapText');
            font.load('bold');
            return { format, font };
          }),
        );
        await ctx.sync();
        if (
          !same(
            [used.rowIndex, used.columnIndex, used.rowCount, used.columnCount],
            [0, 0, 12, 6],
          ) ||
          !same(
            [otherUsed.rowIndex, otherUsed.columnIndex, otherUsed.rowCount, otherUsed.columnCount],
            [0, 0, 1, 3],
          ) ||
          other.formulas.some((row) => row.some((value) => formula(value)))
        )
          throw new Error('样本使用范围或第2表公式不匹配。');
        const measured = widths.map((format) => format.columnWidth);
        if (measured.some((width) => !Number.isFinite(width) || width <= 0))
          unsupported('原生列宽不可靠。');
        return {
          path,
          values: matrix(range.values, 12, 6),
          formulas: range.formulas.map((row) => row.map(formula)),
          second: matrix(other.values, 1, 3),
          catalogue: sheets.items.map((s) => s.name),
          widths: measured,
          bold: formats.map((row) => row.map(({ font }) => bool(font.bold))),
          wrap: formats.map((row) => row.map(({ format }) => bool(format.wrapText))),
          hidden: hidden.map((row) => bool(row.rowHidden)),
        };
      });
    };
    return {
      read,
      notice: async () =>
        unsupported('Office 原任务窗格无法独立枚举并读回新窗口；需人工核对独立文档和完整通知。'),
    };
  }
  const app = getWpsApplication(config.kind);
  const source = app.ActiveDocument;
  const readText = (doc: WpsDocument | undefined): ExtendedNativeState => {
    if (!doc || typeof doc.Content?.Text !== 'string' || !doc.FullName)
      return unsupported('无法读取专用正文和原生绝对路径。');
    assertSourcePath(doc.FullName, config);
    return { path: doc.FullName, text: normalizeText(doc.Content.Text) };
  };
  const read = async (): Promise<ExtendedNativeState> => {
    const active = getWpsApplication(config.kind);
    if (config.kind === 'text') return readText(active.ActiveDocument);
    const book = active.ActiveWorkbook;
    const sheets = book?.Sheets;
    if (!book?.FullName || !sheets || sheets.Count !== 2 || typeof sheets.Item !== 'function')
      return unsupported('WPS 缺少完整原生工作表目录或路径。');
    assertSourcePath(book.FullName, config);
    const catalogue = [sheets.Item(1).Name, sheets.Item(2).Name];
    if (!same(catalogue, ['ExtendedData', 'ScopeSentinel']))
      throw new Error('原生工作表目录不匹配。');
    const sheet = book.Worksheets?.Item('ExtendedData') as ProbeSheet | undefined;
    const second = book.Worksheets?.Item('ScopeSentinel') as ProbeSheet | undefined;
    if (!sheet?.Range || !second?.Range || !sheet.UsedRange || !second.UsedRange)
      return unsupported('WPS 无法可靠读取完整样本范围。');
    const dimensions = (range: ProbeRange) => [
      range.Row,
      range.Column,
      range.Rows?.Count,
      range.Columns?.Count,
    ];
    if (
      !same(dimensions(sheet.UsedRange), [1, 1, 12, 6]) ||
      !same(dimensions(second.UsedRange), [1, 1, 1, 3])
    )
      throw new Error('原生使用范围不匹配。');
    const range = sheet.Range('A1:F12');
    const cells = Array.from({ length: 12 }, (_, r) =>
      Array.from({ length: 6 }, (_, c) => {
        const cell = range.Cells?.Item(r + 1, c + 1) as ProbeRange | undefined;
        if (!cell) return unsupported('WPS 缺少逐格读取接口。');
        return cell;
      }),
    );
    const widths = cells[0]!.map((cell) => cell.ColumnWidth);
    if (widths.some((width) => typeof width !== 'number' || !Number.isFinite(width) || width <= 0))
      return unsupported('WPS 原生列宽不可靠。');
    if (cells.flat().some((cell) => cell.Formula === undefined))
      return unsupported('WPS 无法读取逐格公式，不能仅靠值声称未改公式。');
    const other = second.Range('A1:C1');
    for (let c = 1; c <= 3; c++) {
      const value = (other.Cells?.Item(1, c) as ProbeRange | undefined)?.Formula;
      if (value === undefined || formula(value)) return unsupported('无法核对第2表无公式。');
    }
    return {
      path: book.FullName,
      values: matrix(range.Value2, 12, 6),
      formulas: cells.map((row) => row.map((cell) => formula(cell.Formula))),
      second: matrix(other.Value2, 1, 3),
      catalogue,
      widths: widths as number[],
      bold: cells.map((row) => row.map((cell) => bool(cell.Font?.Bold))),
      wrap: cells.map((row) => row.map((cell) => bool(cell.WrapText))),
      hidden: Array.from({ length: 12 }, (_, r) =>
        bool(sheet.Range!(`A${r + 1}`).EntireRow?.Hidden),
      ),
    };
  };
  return {
    read,
    notice: async () => {
      const current = getWpsApplication('text').ActiveDocument;
      if (
        !current ||
        !current.Content ||
        !source?.FullName ||
        !current.Name ||
        typeof current.Content.Text !== 'string'
      )
        return unsupported('无法读回新建文档。');
      return {
        source: readText(source),
        text: normalizeText(current.Content.Text),
        independent:
          current !== source &&
          (current.FullName || current.Name) !== source.FullName &&
          current.Saved === false,
      };
    },
  };
}

interface ProbeRange extends Omit<WpsSelection, 'Font'> {
  Formula?: unknown;
  ColumnWidth?: number;
  WrapText?: unknown;
  Font?: { Bold?: number | boolean };
  EntireRow?: { Hidden?: unknown };
}
interface ProbeSheet {
  UsedRange?: ProbeRange;
  Range?(address: string): ProbeRange;
}
