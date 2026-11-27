import type {
  CellValue,
  ChangeReceipt,
  DocumentKind,
  ExtendedActionPlan,
  PrintSettingsSnapshot,
  SelectionReadOptions,
  SelectionSnapshot,
  SheetLayoutSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { noticeText, validateExtendedPlan } from '../../core/extended-actions';
import { assertSameSelection, MAX_SELECTION_CELLS } from '../../core/selection';
import { readWpsSheetDetails } from './sheet-details';
import type {
  WpsApplication,
  WpsFileSystem,
  WpsPageSetup,
  WpsSelection,
  WpsSheetPdfSource,
  WpsSheetCell,
  WpsWorkbook,
  WpsWorksheet,
  WpsWriterPdfSelection,
} from './bridge';

// Desktop JSAPI only. No async web-office API or guessed enum values are used.
const unavailable = (message = '当前 WPS 未提供可核对的完整接口，请使用 WPS 原生功能。') =>
  new HelperError('UNAVAILABLE', message);
const stale = () => new HelperError('STALE_SELECTION', '范围或内容已改变，请重新预览。');
function same(a: unknown, b: unknown): boolean {
  if (Object.is(a, b)) return true;
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object') return false;
  if (Array.isArray(a) || Array.isArray(b))
    return (
      Array.isArray(a) &&
      Array.isArray(b) &&
      a.length === b.length &&
      a.every((value, i) => same(value, b[i]))
    );
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  const keys = Object.keys(left);
  return (
    keys.length === Object.keys(right).length &&
    keys.every(
      (key) => Object.prototype.hasOwnProperty.call(right, key) && same(left[key], right[key]),
    )
  );
}
const identity = (object: { Name: string; FullName?: string }) => object.FullName || object.Name;

/** Windows paths only; reject traversal, URL/device schemes and ambiguous trailing characters. */
function absolutePath(raw: string): string | undefined {
  const path = raw.replace(/\//gu, '\\');
  if (path !== path.trim() || /[\u0000-\u001f<>"|?*]/u.test(path) || /^\\\\[?.]\\/u.test(path))
    return;
  const drive = /^[A-Za-z]:\\/u.test(path);
  const unc = /^\\\\[^\\]+\\[^\\]+(?:\\|$)/u.test(path);
  if (!drive && !unc) return;
  if (drive && path.length === 3) return path;
  const rest = path.slice(drive ? 3 : 2);
  if (rest.includes(':')) return;
  const segments = rest.replace(/\\$/u, '').split('\\');
  if (
    segments.some(
      (part) =>
        !part ||
        part === '.' ||
        part === '..' ||
        /[. ]$/u.test(part) ||
        /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/iu.test(part),
    )
  )
    return;
  return path.replace(/\\$/u, '') || path;
}

function pathParent(path: string): string {
  const i = path.lastIndexOf('\\');
  return i === 2 ? path.slice(0, 3) : path.slice(0, i);
}

function exists(fs: WpsFileSystem, path: string): boolean {
  if (typeof fs.Exists !== 'function') throw unavailable('当前版本缺少文件存在性检查。');
  return nativeBoolean(fs.Exists(path));
}

function directoryExists(fs: WpsFileSystem, path: string): boolean {
  try {
    return (
      exists(fs, path) &&
      typeof fs.readdirSync === 'function' &&
      Array.isArray(fs.readdirSync(path))
    );
  } catch {
    return false;
  }
}

/** Sheets (not Worksheets) covers both worksheets and chart sheets. Never activate/unhide. */
function pdfSheetScope(app: WpsApplication, book: WpsWorkbook): string[] {
  const visible = app.Enum?.xlSheetVisible ?? app.Enum?.XlSheetVisibility?.xlSheetVisible;
  if (!Number.isSafeInteger(visible)) throw unavailable('无法读取工作表可见状态枚举。');
  const sheets = book.Sheets;
  const count = sheets?.Count;
  // This is an assistant safety budget, not a native worksheet-count limit.
  if (
    !sheets ||
    typeof sheets.Item !== 'function' ||
    !Number.isSafeInteger(count) ||
    !count ||
    count < 1 ||
    count > 1000
  )
    throw unavailable('无法完整枚举所有工作表（含图表表），或工作表数量超出 1,000 张核对范围。');
  const names: string[] = [];
  const seen = new Set<string>();
  for (let index = 1; index <= count; index++) {
    const sheet = sheets.Item(index);
    if (
      !sheet ||
      typeof sheet.Name !== 'string' ||
      !sheet.Name.trim() ||
      !sheet.Parent ||
      identity(sheet.Parent) !== identity(book) ||
      seen.has(sheet.Name.toLowerCase())
    )
      throw unavailable('所有工作表的目录、归属或名称无法完整核对。');
    if (sheet.Visible !== visible)
      throw unavailable('存在隐藏工作表或未知可见状态，无法保证所有工作表均被导出。');
    names.push(sheet.Name);
    seen.add(sheet.Name.toLowerCase());
  }
  if (sheets.Count !== count) throw unavailable('工作表目录在核对期间发生变化，请重新预览。');
  return names;
}

/** Grouped worksheets can change native export scope; never ungroup them for the user. */
function assertSinglePdfSheet(app: WpsApplication, book: WpsWorkbook, sheet: WpsWorksheet): void {
  const selected = app.ActiveWindow?.SelectedSheets;
  if (!selected || selected.Count !== 1 || typeof selected.Item !== 'function')
    throw unavailable(
      '无法确认仅选中一张工作表，或当前工作表已成组选择。请在原生导出窗口核对范围。',
    );
  const current = selected.Item(1);
  if (
    !current ||
    current.Name !== sheet.Name ||
    app.ActiveSheet?.Name !== sheet.Name ||
    !current.Parent ||
    !sheet.Parent ||
    !app.ActiveSheet.Parent ||
    identity(current.Parent) !== identity(book) ||
    identity(sheet.Parent) !== identity(book) ||
    identity(app.ActiveSheet.Parent) !== identity(book) ||
    selected.Count !== 1
  )
    throw unavailable('当前工作表的单表选择状态或归属无法可靠确认，请在原生导出窗口核对范围。');
}

function assertPdfTextBody(
  app: WpsApplication,
  selection: WpsSelection | undefined,
  expected: SelectionSnapshot,
  documentId: string,
): void {
  const main = app.Enum?.wdMainTextStory ?? app.Enum?.WdStoryType?.wdMainTextStory;
  const range = selection?.Range;
  if (
    !Number.isSafeInteger(main) ||
    !range ||
    range.StoryType !== main ||
    !range.Document ||
    identity(range.Document) !== documentId ||
    `${selection?.Start}:${selection?.End}` !== expected.targetId ||
    `${range.Start}:${range.End}` !== expected.targetId ||
    selection?.Text !== expected.text ||
    range.Text !== expected.text
  )
    throw unavailable(
      '选区 PDF 仅支持可准确确认的正文选区；页眉、页脚、文本框及未知文字区域请使用原生导出。',
    );
}

/** ReadFile variants are accepted only when their byte representation is unambiguous. */
const MAX_FILE_BYTES = 128 * 1024 * 1024;

function fileBytes(fs: WpsFileSystem, path: string): string {
  const raw =
    typeof fs.readAsBinaryString === 'function' ? fs.readAsBinaryString(path) : fs.ReadFile?.(path);
  if (typeof raw === 'string') {
    if (raw.length > MAX_FILE_BYTES)
      throw unavailable(
        '文件超过 128 MiB，助手不继续读取、比较或报告保存成功。请使用 WPS 原生功能核验。',
      );
    if (raw.length > 0 && !/[^\u0000-\u00ff]/u.test(raw)) return raw;
  }
  const bytes =
    raw instanceof Uint8Array ? raw : raw instanceof ArrayBuffer ? new Uint8Array(raw) : undefined;
  if (!bytes?.length) throw unavailable('文件内容无法按字节读回核验。');
  if (bytes.length > MAX_FILE_BYTES)
    throw unavailable('文件超过 128 MiB，助手不继续转换或比较。请使用 WPS 原生功能核验。');
  let result = '';
  for (let i = 0; i < bytes.length; i += 8192)
    result += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return result;
}

function nativeBoolean(value: unknown): boolean {
  if (value === true || value === -1) return true;
  if (value === false || value === 0) return false;
  throw unavailable('无法读取单个单元格的明确格式或合并状态，请使用 WPS 原生功能。');
}

function dimensions(range: WpsSelection): [number, number] {
  const rows = range.Rows?.Count;
  const columns = range.Columns?.Count;
  if (
    range.Areas?.Count !== 1 ||
    !Number.isInteger(rows) ||
    !Number.isInteger(columns) ||
    !rows ||
    !columns ||
    rows < 1 ||
    columns < 1
  )
    throw unavailable('只支持可完整读取的单个连续矩形区域，不支持多区域选区。');
  if (rows * columns > MAX_SELECTION_CELLS)
    throw new HelperError('SELECTION_TOO_LARGE', '请缩小到 5,000 个单元格以内。');
  return [rows, columns];
}

export function readWpsValues(range: WpsSelection, rows: number, columns: number): CellValue[][] {
  const raw = range.Value2;
  const matrix = Array.isArray(raw) ? raw : rows === 1 && columns === 1 ? [[raw]] : undefined;
  if (
    !matrix ||
    matrix.length !== rows ||
    !matrix.every(
      (row) =>
        Array.isArray(row) &&
        row.length === columns &&
        row.every(
          (value) =>
            value === null ||
            typeof value === 'string' ||
            typeof value === 'boolean' ||
            typeof value === 'number',
        ),
    )
  )
    throw unavailable('WPS 返回的数据不完整或含有无法识别的值，无法安全处理。');
  return matrix.map((row) => [...row]) as CellValue[][];
}

function addressOf(range: WpsSelection): string {
  const address = typeof range.Address === 'function' ? range.Address() : range.Address;
  if (
    typeof address !== 'string' ||
    !/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/iu.test(address)
  )
    throw unavailable('无法确认单个连续区域的地址，请重新选择单元格。');
  return address;
}

function cellAt(range: WpsSelection, row: number, column: number): WpsSheetCell {
  if (typeof range.Cells?.Item !== 'function') throw unavailable();
  const cell = range.Cells.Item(row, column);
  if (!cell || typeof cell !== 'object') throw unavailable();
  return cell;
}

function sheetOf(app: WpsApplication, range: WpsSelection): WpsWorksheet {
  const book = app.ActiveWorkbook;
  const sheet = range.Worksheet;
  if (
    !book ||
    !sheet ||
    !sheet.Parent ||
    !app.ActiveSheet ||
    sheet.Name !== app.ActiveSheet.Name ||
    identity(sheet.Parent) !== identity(book)
  )
    throw unavailable('无法确认选区属于当前工作簿与工作表。');
  if (typeof book.ReadOnly !== 'boolean' || typeof sheet.ProtectContents !== 'boolean')
    throw unavailable('无法读取工作簿只读状态或工作表保护状态。');
  return sheet;
}

/** ColumnWidth is the native character-width unit, not a point measurement. */
export function readWpsSheetLayout(app: WpsApplication, range: WpsSelection): SheetLayoutSnapshot {
  const [rows, columns] = dimensions(range);
  const sheet = sheetOf(app, range);
  if (
    !Number.isInteger(range.Row) ||
    !Number.isInteger(range.Column) ||
    !range.Row ||
    !range.Column ||
    range.Row < 1 ||
    range.Column < 1
  )
    throw unavailable();
  const columnWidths: number[] = [];
  const headerBold: boolean[] = [];
  const wrapText: boolean[][] = [];
  let merged = false;
  for (let r = 1; r <= rows; r++) {
    const wraps: boolean[] = [];
    for (let c = 1; c <= columns; c++) {
      const cell = cellAt(range, r, c);
      merged = nativeBoolean(cell.MergeCells) || merged;
      wraps.push(nativeBoolean(cell.WrapText));
      if (r === 1) {
        const width = cell.ColumnWidth;
        if (typeof width !== 'number' || !Number.isFinite(width) || width < 0 || width > 255)
          throw unavailable('无法准确读取各列原始列宽。');
        columnWidths.push(width);
        headerBold.push(nativeBoolean(cell.Font?.Bold));
      }
    }
    wrapText.push(wraps);
  }
  return {
    columnWidths,
    widthUnit: 'characters',
    wrapText,
    headerBold,
    merged,
    protected: app.ActiveWorkbook!.ReadOnly! || sheet.ProtectContents!,
    firstRow: range.Row,
    firstColumn: range.Column,
    sheetName: sheet.Name,
  };
}

type PrintState = {
  PrintArea: string;
  Orientation: number;
  Zoom: number | false;
  FitToPagesWide: number | false;
  FitToPagesTall: number | false;
};
const printKeys = ['PrintArea', 'Orientation', 'Zoom', 'FitToPagesWide', 'FitToPagesTall'] as const;

function orientationEnum(app: WpsApplication, key: 'xlLandscape' | 'xlPortrait'): number {
  const value = app.Enum?.[key] ?? app.Enum?.XlPageOrientation?.[key];
  if (typeof value !== 'number' || !Number.isFinite(value))
    throw unavailable('当前版本缺少页面方向枚举。');
  return value;
}

function printState(setup: WpsPageSetup | undefined): PrintState {
  if (
    !setup ||
    typeof setup.PrintArea !== 'string' ||
    typeof setup.Orientation !== 'number' ||
    !Number.isFinite(setup.Orientation) ||
    !(
      setup.Zoom === false ||
      (typeof setup.Zoom === 'number' && Number.isFinite(setup.Zoom) && setup.Zoom > 0)
    ) ||
    ![setup.FitToPagesWide, setup.FitToPagesTall].every(
      (value) =>
        value === false || (typeof value === 'number' && Number.isInteger(value) && value >= 0),
    )
  )
    throw unavailable('当前版本无法完整读取页面设置，不能安全写入或恢复。');
  return {
    PrintArea: setup.PrintArea,
    Orientation: setup.Orientation,
    Zoom: setup.Zoom,
    FitToPagesWide: setup.FitToPagesWide!,
    FitToPagesTall: setup.FitToPagesTall!,
  };
}

export function readWpsPrintSettings(
  app: WpsApplication,
  range: WpsSelection,
): PrintSettingsSnapshot {
  try {
    dimensions(range);
    const sheet = sheetOf(app, range);
    addressOf(range);
    if (typeof sheet.Range !== 'function')
      throw unavailable('当前版本无法取得可核对的固定选区，不能安全设置或恢复打印参数。');
    const state = printState(sheet.PageSetup);
    const landscape = orientationEnum(app, 'xlLandscape');
    const portrait = orientationEnum(app, 'xlPortrait');
    if (landscape === portrait || ![landscape, portrait].includes(state.Orientation))
      throw unavailable();
    return {
      available: true,
      printArea: state.PrintArea,
      orientation: state.Orientation === landscape ? 'landscape' : 'portrait',
      zoom: state.Zoom,
      fitWide: state.FitToPagesWide === false ? 0 : state.FitToPagesWide,
      fitTall: state.FitToPagesTall === false ? 0 : state.FitToPagesTall,
    };
  } catch (error) {
    return {
      available: false,
      reason: error instanceof HelperError ? error.message : 'WPS 页面设置读取失败。',
      printArea: '',
      orientation: '',
      zoom: false,
      fitWide: 0,
      fitTall: 0,
    };
  }
}

function readOptions(expected: SelectionSnapshot): SelectionReadOptions {
  return {
    includeCellTypes: expected.cellTypes !== undefined,
    includeSheetLayout: expected.sheetLayout !== undefined,
    includePrintSettings: expected.printSettings !== undefined,
  };
}

function assertEditable(layout: SheetLayoutSnapshot): void {
  if (layout.merged || layout.protected)
    throw unavailable('合并单元格、受保护工作表或只读工作簿不能由助手修改。');
}

function layoutProperty(
  type: ExtendedActionPlan['type'],
): 'columnWidths' | 'wrapText' | 'headerBold' {
  if (type === 'sheet.autofit') return 'columnWidths';
  return type === 'sheet.wrap' ? 'wrapText' : 'headerBold';
}

type SheetChange = {
  token: string;
  documentId: string;
  range: WpsSelection;
  address: string;
  values: CellValue[][];
  before: SheetLayoutSnapshot;
  after: SheetLayoutSnapshot;
} & (
  | { type: 'sheet.autofit' | 'sheet.wrap' | 'sheet.header' }
  | { type: 'sheet.print'; setup: WpsPageSetup; beforePrint: PrintState; afterPrint: PrintState }
);

interface TaskHost {
  kind: DocumentKind;
  application(): WpsApplication;
  documentId(app: WpsApplication): string;
  readSelection(options?: SelectionReadOptions): Promise<SelectionSnapshot>;
  invalidateText(): void;
}

export class WpsTasks {
  private change?: SheetChange;
  private readonly results = new Map<string, { path: string; file?: string }>();
  constructor(private readonly host: TaskHost) {}

  clear(): void {
    this.change = undefined;
  }
  dispose(): void {
    this.clear();
    this.results.clear();
  }
  owns(receipt: ChangeReceipt): boolean {
    return receipt.token.startsWith('wps-task:');
  }

  private guidance(message: string, outcome: 'guidance' | 'cancelled' = 'guidance'): ChangeReceipt {
    return { token: `wps-task:${crypto.randomUUID()}`, canUndo: false, outcome, message };
  }

  async apply(plan: ExtendedActionPlan, expected: SelectionSnapshot): Promise<ChangeReceipt> {
    validateExtendedPlan(plan);
    assertSameSelection(expected, await this.host.readSelection(readOptions(expected)));
    const app = this.host.application();
    if (this.host.documentId(app) !== expected.documentId) throw stale();
    if (plan.type === 'document.saveCopy' || plan.type === 'document.exportPdf')
      return this.save(app, plan, expected);
    if (plan.type === 'document.notice') return this.notice(app, plan);
    if (this.host.kind !== 'sheet' || expected.kind !== 'sheet')
      throw unavailable('请在 WPS 表格中操作。');
    return this.sheet(app, plan, expected);
  }

  private fixedRange(app: WpsApplication, expected: SelectionSnapshot): WpsSelection {
    const selection = app.Selection;
    if (!selection) throw stale();
    const sheet = sheetOf(app, selection);
    const address = addressOf(selection);
    if (`${sheet.Name}!${address}` !== expected.targetId || typeof sheet.Range !== 'function')
      throw stale();
    const range = sheet.Range(address);
    const [rows, columns] = dimensions(range);
    if (
      addressOf(range) !== address ||
      sheetOf(app, range).Name !== sheet.Name ||
      !same(readWpsValues(range, rows, columns), expected.values)
    )
      throw stale();
    if (expected.cellTypes) {
      const details = readWpsSheetDetails(range, app, expected.values!);
      if (!same(details.cellTypes, expected.cellTypes) || details.rowScope !== expected.rowScope)
        throw stale();
    }
    if (expected.sheetLayout && !same(readWpsSheetLayout(app, range), expected.sheetLayout))
      throw stale();
    if (expected.printSettings && !same(readWpsPrintSettings(app, range), expected.printSettings))
      throw stale();
    return range;
  }

  private sheet(
    app: WpsApplication,
    plan: Exclude<
      ExtendedActionPlan,
      { type: 'document.saveCopy' | 'document.exportPdf' } | { type: 'document.notice' }
    >,
    expected: SelectionSnapshot,
  ): ChangeReceipt {
    // A preview without the original opt-in layout cannot authorize a format write.
    if (!expected.sheetLayout) {
      if (plan.type === 'sheet.print')
        return this.guidance(
          '尚未取得完整选区版式快照，未更改打印设置。请在 WPS 页面布局中设置当前选区、横向、宽 1 页和高自动，并打开原生打印预览。',
        );
      throw unavailable('请重新预览以读取原始列宽、换行、表头和保护状态。');
    }
    if (plan.type === 'sheet.print' && !expected.printSettings?.available)
      return this.guidance(
        `${expected.printSettings?.reason ?? '当前版本没有完整页面设置接口。'}未更改打印设置。请在 WPS 中设置打印区域为当前选区、横向、宽 1 页和高自动，再打开原生打印预览。`,
      );
    const range = this.fixedRange(app, expected);
    if (
      expected.values?.some((row) =>
        row.some((value) => typeof value === 'number' && !Number.isFinite(value)),
      )
    )
      throw unavailable('选区包含非有限数值，不能安全修改格式。');
    const before = readWpsSheetLayout(app, range);
    assertEditable(before);
    if (plan.type === 'sheet.autofit' && typeof range.Columns?.AutoFit !== 'function')
      throw unavailable('当前版本没有可用的列宽自适应接口。');
    const beforePrint =
      plan.type === 'sheet.print' ? printState(range.Worksheet?.PageSetup) : undefined;
    const setup = range.Worksheet?.PageSetup;
    const desiredPrint: PrintState | undefined = beforePrint
      ? {
          PrintArea: addressOf(range),
          Orientation: orientationEnum(app, 'xlLandscape'),
          Zoom: false,
          FitToPagesWide: 1,
          FitToPagesTall: false,
        }
      : undefined;
    this.clear();
    this.host.invalidateText();
    let after: SheetLayoutSnapshot;
    let afterPrint: PrintState | undefined;
    try {
      if (plan.type === 'sheet.print') this.writePrint(setup!, desiredPrint!);
      else this.writeLayout(range, plan.type);
      after = readWpsSheetLayout(app, range);
      const [rows, columns] = dimensions(range);
      if (
        this.host.documentId(app) !== expected.documentId ||
        !same(readWpsValues(range, rows, columns), expected.values)
      )
        throw stale();
      if (plan.type === 'sheet.print') {
        afterPrint = printState(setup);
        if (!same(after, before) || !this.printMatches(afterPrint, desiredPrint!)) throw stale();
      } else {
        const key = layoutProperty(plan.type);
        const projected = { ...before, [key]: after[key] };
        if (
          !same(projected, after) ||
          (plan.type === 'sheet.wrap' &&
            after.wrapText.some((row) => row.some((value) => !value))) ||
          (plan.type === 'sheet.header' && after.headerBold.some((value) => !value))
        )
          throw stale();
      }
    } catch {
      throw new HelperError(
        'INVALID_PLAN',
        'WPS 可能已部分修改，但读回核对未通过。请检查原范围并使用 WPS 原生撤销；助手没有可用的恢复凭证。',
      );
    }
    const token = `wps-task:${crypto.randomUUID()}`;
    const base = {
      token,
      documentId: expected.documentId,
      range,
      address: addressOf(range),
      values: expected.values!.map((row) => [...row]),
      before,
      after,
    };
    this.change =
      plan.type === 'sheet.print'
        ? {
            ...base,
            type: plan.type,
            setup: setup!,
            beforePrint: beforePrint!,
            afterPrint: afterPrint!,
          }
        : { ...base, type: plan.type };
    const messages = {
      'sheet.autofit': '已对选区所涉及的整列自适应列宽，可恢复各列原宽度。',
      'sheet.wrap':
        '已对选区启用自动换行。恢复仅恢复各格原换行属性；WPS 自动调整的行高请在原生界面检查。',
      'sheet.header': '已加粗选区首行，可逐格恢复原粗细。',
      'sheet.print':
        '已设置当前选区为打印区域、横向、宽 1 页和高自动。请在 WPS 按 Ctrl+P 打开原生打印预览，核对纸张、分页及缩放后自行打印；助手没有打印或测算页数。',
    };
    return { token, canUndo: true, outcome: 'completed', message: messages[plan.type] };
  }

  private writeLayout(
    range: WpsSelection,
    type: 'sheet.autofit' | 'sheet.wrap' | 'sheet.header',
    restore?: SheetLayoutSnapshot,
  ): void {
    const [rows, columns] = dimensions(range);
    if (type === 'sheet.autofit' && !restore) {
      range.Columns!.AutoFit!();
      return;
    }
    for (let r = 1; r <= (type === 'sheet.wrap' ? rows : 1); r++)
      for (let c = 1; c <= columns; c++) {
        const cell = cellAt(range, r, c);
        if (type === 'sheet.autofit') cell.ColumnWidth = restore!.columnWidths[c - 1]!;
        else if (type === 'sheet.wrap')
          cell.WrapText = restore ? restore.wrapText[r - 1]![c - 1]! : true;
        else cell.Font!.Bold = (restore ? restore.headerBold[c - 1]! : true) ? -1 : 0;
      }
  }

  /** Restore only cells/columns actually changed by this receipt, preserving independent edits. */
  private layoutRestoration(
    change: SheetChange,
    current: SheetLayoutSnapshot,
  ): SheetLayoutSnapshot {
    const result = structuredClone(current);
    if (change.type === 'sheet.print') return result;
    const property = layoutProperty(change.type);
    const restore = (before: unknown, after: unknown, actual: unknown) => {
      if (same(before, after)) return actual;
      if (!same(after, actual))
        throw new HelperError('STALE_SELECTION', '要恢复的格式后来又被修改，不能覆盖新的修改。');
      return before;
    };
    if (property === 'wrapText') {
      result.wrapText = current.wrapText.map((row, r) =>
        row.map(
          (value, c) =>
            restore(change.before.wrapText[r]![c], change.after.wrapText[r]![c], value) as boolean,
        ),
      );
    } else if (property === 'columnWidths') {
      result.columnWidths = current.columnWidths.map(
        (value, c) =>
          restore(change.before.columnWidths[c], change.after.columnWidths[c], value) as number,
      );
    } else {
      result.headerBold = current.headerBold.map(
        (value, c) =>
          restore(change.before.headerBold[c], change.after.headerBold[c], value) as boolean,
      );
    }
    return result;
  }

  private restoreLayout(
    range: WpsSelection,
    type: 'sheet.autofit' | 'sheet.wrap' | 'sheet.header',
    current: SheetLayoutSnapshot,
    desired: SheetLayoutSnapshot,
  ): void {
    const [rows, columns] = dimensions(range);
    for (let r = 0; r < (type === 'sheet.wrap' ? rows : 1); r++)
      for (let c = 0; c < columns; c++) {
        const cell = cellAt(range, r + 1, c + 1);
        if (type === 'sheet.autofit' && current.columnWidths[c] !== desired.columnWidths[c])
          cell.ColumnWidth = desired.columnWidths[c]!;
        else if (type === 'sheet.wrap' && current.wrapText[r]![c] !== desired.wrapText[r]![c])
          cell.WrapText = desired.wrapText[r]![c]!;
        else if (type === 'sheet.header' && current.headerBold[c] !== desired.headerBold[c])
          cell.Font!.Bold = desired.headerBold[c] ? -1 : 0;
      }
  }

  private writePrint(setup: WpsPageSetup, values: PrintState): void {
    // Zoom=false must precede the fit-to-page values; never call PrintOut.
    setup.PrintArea = values.PrintArea;
    setup.Orientation = values.Orientation;
    setup.Zoom = false;
    setup.FitToPagesWide = values.FitToPagesWide;
    setup.FitToPagesTall = values.FitToPagesTall;
    setup.Zoom = values.Zoom;
  }

  private printMatches(actual: PrintState, expected: PrintState): boolean {
    return printKeys.every((key) =>
      key === 'FitToPagesTall' || key === 'FitToPagesWide'
        ? (actual[key] || 0) === (expected[key] || 0)
        : actual[key] === expected[key],
    );
  }

  async undo(receipt: ChangeReceipt): Promise<void> {
    const change = this.change;
    if (!change || change.token !== receipt.token)
      throw new HelperError('EXPIRED_TASK', '这次修改已无法从助手恢复。');
    const app = this.host.application();
    if (this.host.documentId(app) !== change.documentId) throw stale();
    // The stored native Range remains tied to the original cells after selection moves.
    const sheet = change.range.Worksheet;
    if (
      !sheet ||
      !sheet.Parent ||
      !app.ActiveWorkbook ||
      identity(sheet.Parent) !== identity(app.ActiveWorkbook)
    )
      throw stale();
    // Reading another sheet's fixed range must not activate or redirect the user's selection.
    if (app.ActiveSheet?.Name !== sheet.Name)
      throw new HelperError('STALE_SELECTION', '请回到原工作表后再恢复。');
    const current = readWpsSheetLayout(app, change.range);
    assertEditable(current);
    const [rows, columns] = dimensions(change.range);
    if (
      addressOf(change.range) !== change.address ||
      !same(readWpsValues(change.range, rows, columns), change.values) ||
      !same(
        [current.firstRow, current.firstColumn, current.sheetName],
        [change.after.firstRow, change.after.firstColumn, change.after.sheetName],
      )
    )
      throw stale();
    let desiredPrint: PrintState | undefined;
    let printRestoreKeys: (typeof printKeys)[number][] = [];
    const projected = this.layoutRestoration(change, current);
    if (change.type === 'sheet.print') {
      const currentPrint = printState(change.setup);
      desiredPrint = { ...currentPrint };
      printRestoreKeys = printKeys.filter(
        (key) => !Object.is(change.beforePrint[key], change.afterPrint[key]),
      );
      for (const key of printRestoreKeys) {
        if (!Object.is(currentPrint[key], change.afterPrint[key]))
          throw new HelperError(
            'STALE_SELECTION',
            '要恢复的打印属性后来又被修改，不能覆盖新设置。',
          );
      }
      desiredPrint = {
        ...currentPrint,
        ...Object.fromEntries(printRestoreKeys.map((key) => [key, change.beforePrint[key]])),
      };
    }
    try {
      if (change.type === 'sheet.print') {
        // Assign only changed properties; keep subsequent user edits to untouched settings.
        for (const key of printRestoreKeys) {
          if (key === 'PrintArea') change.setup.PrintArea = desiredPrint!.PrintArea;
          else if (key === 'Orientation') change.setup.Orientation = desiredPrint!.Orientation;
          else if (key === 'Zoom') change.setup.Zoom = desiredPrint!.Zoom;
          else if (key === 'FitToPagesWide')
            change.setup.FitToPagesWide = desiredPrint!.FitToPagesWide;
          else change.setup.FitToPagesTall = desiredPrint!.FitToPagesTall;
        }
      } else this.restoreLayout(change.range, change.type, current, projected);
      const after = readWpsSheetLayout(app, change.range);
      if (
        !same(after, projected) ||
        !same(readWpsValues(change.range, rows, columns), change.values) ||
        (change.type === 'sheet.print' && !same(printState(change.setup), desiredPrint))
      )
        throw stale();
    } catch {
      this.clear();
      throw new HelperError(
        'INVALID_PLAN',
        '恢复可能只完成了一部分，请查看原范围并使用 WPS 原生撤销。',
      );
    }
    this.clear();
  }

  private async save(
    app: WpsApplication,
    plan: Extract<ExtendedActionPlan, { type: 'document.saveCopy' | 'document.exportPdf' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    const label = plan.type === 'document.exportPdf' ? '导出 PDF' : '保存副本';
    const scope = plan.type === 'document.exportPdf' ? (plan.scope ?? 'document') : 'document';
    const scopeLabel =
      scope === 'selection'
        ? '仅当前选区'
        : scope === 'activeSheet'
          ? '仅当前工作表（忽略原打印区域）'
          : this.host.kind === 'sheet'
            ? '整个工作簿（所有工作表）'
            : '整份文档';
    const next =
      plan.type === 'document.exportPdf'
        ? '在 WPS 的“输出为 PDF”中选择此位置并确认导出'
        : '在 WPS 的“另存为／保存副本”中选择此位置并确认保存';
    const hint = `请${next}，文件名使用“${plan.fileName}”${plan.folder ? `，文件夹为“${plan.folder}”` : ''}${plan.type === 'document.exportPdf' ? `，范围选择“${scopeLabel}”` : ''}；完成后核对实际文件。助手尚未${label}，不会覆盖原件。`;
    const pdf = plan.type === 'document.exportPdf';
    const pdfWorkbook = pdf && this.host.kind === 'sheet' ? app.ActiveWorkbook : undefined;
    const source = this.host.kind === 'sheet' ? app.ActiveWorkbook : app.ActiveDocument;
    const fs = app.FileSystem;
    let sheetScope: string[] | undefined;
    let sheetExporter: WpsSheetPdfSource | undefined;
    let writerSelection: WpsWriterPdfSelection | undefined;
    if (pdf && scope === 'activeSheet' && this.host.kind !== 'sheet')
      return this.guidance('当前工作表范围仅适用于表格；助手未导出或保存。');
    if (pdf && scope === 'selection' && this.host.kind === 'text') {
      writerSelection = app.Selection as WpsWriterPdfSelection | undefined;
      if (!expected.text?.trim() || typeof writerSelection?.ExportAsFixedFormat !== 'function')
        return this.guidance(`当前宿主没有可靠的文字选区导出接口。${hint}`);
      try {
        if (!source) throw unavailable('无法确认原稿。');
        assertPdfTextBody(app, writerSelection, expected, identity(source));
      } catch {
        return this.guidance(`选区 PDF 仅支持可确认的正文区域；当前文字区域无法可靠核对。${hint}`);
      }
    }
    if (pdf && this.host.kind === 'sheet') {
      try {
        if (!pdfWorkbook) throw unavailable('无法取得当前工作簿。');
        if (scope === 'document') {
          sheetScope = pdfSheetScope(app, pdfWorkbook);
          sheetExporter = pdfWorkbook;
        } else {
          const range = this.fixedRange(app, expected);
          assertSinglePdfSheet(app, pdfWorkbook, sheetOf(app, range));
          sheetExporter =
            scope === 'selection'
              ? (range as WpsSelection & WpsSheetPdfSource)
              : sheetOf(app, range);
        }
        if (typeof sheetExporter?.ExportAsFixedFormat !== 'function')
          throw unavailable('当前宿主没有所选范围的 PDF 导出接口。');
        if (typeof fs?.readdirSync !== 'function') throw unavailable('无法可靠核对目标文件夹。');
      } catch (error) {
        const reason = error instanceof HelperError ? error.message : 'WPS 工作表目录读取失败。';
        return this.guidance(
          `${reason}无法可靠确认 PDF 范围为${scopeLabel}，助手未自动导出或保存。${scope === 'document' ? '请在 WPS“输出为 PDF”中核对所有工作表（包括隐藏工作表）均纳入导出，并检查各表打印区域没有遗漏内容，确认完整预览后再导出。' : ''}${hint}`,
        );
      }
    }
    if (
      !source ||
      !fs ||
      typeof fs.Exists !== 'function' ||
      (typeof fs.readAsBinaryString !== 'function' && typeof fs.ReadFile !== 'function')
    )
      return this.guidance(hint);
    const sourceId = identity(source);
    const extension = /\.[^.]+$/u.exec(source.Name)?.[0]?.toLowerCase();
    if (
      !pdf &&
      (!extension ||
        !/\.(xlsx?|xlsm|xlsb|xltx|xltm|et|ett|csv|docx?|docm|dotx|dotm|wps|wpt)$/iu.test(
          extension,
        ) ||
        !plan.fileName.toLowerCase().endsWith(extension))
    )
      return this.guidance(`保存副本必须保留原文件格式及扩展名，助手未保存。${hint}`);
    const pdfEnum = pdfWorkbook
      ? (app.Enum?.xlTypePDF ?? app.Enum?.XlFixedFormatType?.xlTypePDF)
      : (app.Enum?.wdExportFormatPDF ?? app.Enum?.WdExportFormat?.wdExportFormatPDF);
    if (
      pdf
        ? (!sheetExporter &&
            !writerSelection &&
            typeof source.ExportAsFixedFormat !== 'function') ||
          !Number.isSafeInteger(pdfEnum)
        : this.host.kind === 'sheet'
          ? typeof app.ActiveWorkbook?.SaveCopyAs !== 'function'
          : typeof fs.copyFileSync !== 'function' ||
            app.ActiveDocument?.Saved !== true ||
            !app.ActiveDocument.FullName
    )
      return this.guidance(
        `${!pdf && this.host.kind === 'text' ? '文字副本仅支持已保存且没有未保存改动的磁盘原稿。' : ''}${hint}`,
      );
    let target = plan.folder
      ? absolutePath(`${plan.folder.replace(/[\\/]+$/u, '')}\\${plan.fileName}`)
      : undefined;
    if (plan.folder && !target)
      return this.guidance(`目标必须是明确的 Windows 绝对路径，助手未保存。${hint}`);
    let attempted = false;
    try {
      if (!target) {
        const kind =
          app.Enum?.msoFileDialogSaveAs ?? app.Enum?.MsoFileDialogType?.msoFileDialogSaveAs;
        if (typeof app.FileDialog !== 'function' || typeof kind !== 'number')
          return this.guidance(hint);
        const dialog = app.FileDialog(kind);
        if (
          typeof dialog?.Show !== 'function' ||
          !('InitialFileName' in dialog) ||
          !('Title' in dialog) ||
          !('ButtonName' in dialog)
        )
          return this.guidance(hint);
        dialog.InitialFileName = plan.fileName;
        dialog.Title = `${label}：选择新文件位置，已有文件不会覆盖`;
        dialog.ButtonName = '选择位置';
        const result = dialog.Show();
        if (result === 0)
          return this.guidance(`已取消${label}；助手未保存或导出文件。`, 'cancelled');
        if (
          result !== -1 ||
          dialog.SelectedItems?.Count !== 1 ||
          typeof dialog.SelectedItems.Item !== 'function'
        )
          return this.guidance(hint);
        const selected = dialog.SelectedItems.Item(1);
        if (typeof selected !== 'string') return this.guidance(hint);
        target = absolutePath(selected);
        if (
          !target ||
          target.slice(target.lastIndexOf('\\') + 1).toLowerCase() !== plan.fileName.toLowerCase()
        )
          return this.guidance(
            `路径或文件名与预览不一致，请重新填写文件名后预览。助手未保存。${hint}`,
          );
      }
      assertSameSelection(expected, await this.host.readSelection(readOptions(expected)));
      if (
        this.host.documentId(this.host.application()) !== expected.documentId ||
        identity(source) !== sourceId
      )
        throw stale();
      const sourcePath = source.FullName ? absolutePath(source.FullName) : undefined;
      if (sourcePath?.toLowerCase() === target.toLowerCase())
        throw new HelperError('INVALID_PLAN', '目标与原文件相同，请更换文件名；原文件没有被覆盖。');
      if (pdfWorkbook ? !directoryExists(fs, pathParent(target)) : !exists(fs, pathParent(target)))
        return this.guidance(`目标文件夹不存在或无法可靠确认为目录，助手未保存。${hint}`);
      if (exists(fs, target))
        throw new HelperError('INVALID_PLAN', '目标已存在，请更换新文件名；助手不会覆盖已有文件。');
      let originalBytes: string | undefined;
      const sourceDocument = this.host.kind === 'text' ? app.ActiveDocument : undefined;
      let originalText: string | undefined;
      if (!pdf && this.host.kind === 'text') {
        if (
          !sourcePath ||
          !exists(fs, sourcePath) ||
          sourceDocument?.Saved !== true ||
          typeof sourceDocument.Content?.Text !== 'string'
        )
          return this.guidance(hint);
        originalText = sourceDocument.Content.Text;
        originalBytes = fileBytes(fs, sourcePath);
      }
      // Recheck immediately before the host's file-creation call. No SaveAs/Execute is called.
      if (exists(fs, target))
        throw new HelperError('INVALID_PLAN', '目标刚刚被创建，请更换文件名。');
      if (
        originalText !== undefined &&
        (sourceDocument?.Saved !== true || sourceDocument.Content?.Text !== originalText)
      )
        throw stale();
      if (sheetScope && !same(pdfSheetScope(this.host.application(), pdfWorkbook!), sheetScope))
        return this.guidance(`所有工作表的目录或可见状态发生变化，助手未导出，请重新预览。${hint}`);
      if (pdf && scope !== 'document') {
        const currentApp = this.host.application();
        assertSameSelection(expected, await this.host.readSelection(readOptions(expected)));
        if (
          this.host.documentId(currentApp) !== expected.documentId ||
          identity(source) !== sourceId
        )
          throw stale();
        if (this.host.kind === 'sheet') {
          const currentRange = this.fixedRange(currentApp, expected);
          const originalSheet =
            scope === 'selection'
              ? sheetOf(currentApp, sheetExporter as WpsSelection)
              : (sheetExporter as WpsWorksheet);
          if (
            originalSheet.Name !== sheetOf(currentApp, currentRange).Name ||
            !originalSheet.Parent ||
            identity(originalSheet.Parent) !== sourceId
          )
            throw stale();
          assertSinglePdfSheet(currentApp, pdfWorkbook!, originalSheet);
        } else {
          assertPdfTextBody(currentApp, writerSelection, expected, sourceId);
          assertPdfTextBody(currentApp, currentApp.Selection, expected, sourceId);
        }
      }
      attempted = true;
      // Official desktop signature: ignore every print area; omit page bounds and
      // disable opening. Whole-book scope is authorized only by the Sheets checks.
      // https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Workbook/member/ExportAsFixedFormat
      if (sheetExporter)
        await sheetExporter.ExportAsFixedFormat!(
          pdfEnum!,
          target,
          undefined,
          undefined,
          true,
          undefined,
          undefined,
          false,
        );
      // ExportCurrentPage=false is essential: true would include unselected text.
      // Official WPS desktop Selection.ExportAsFixedFormat signature.
      else if (writerSelection)
        await writerSelection.ExportAsFixedFormat!(target, pdfEnum!, false, undefined, false);
      else if (pdf) await app.ActiveDocument!.ExportAsFixedFormat!(target, pdfEnum!, false);
      else if (this.host.kind === 'sheet') await app.ActiveWorkbook!.SaveCopyAs!(target);
      else await fs.copyFileSync!(sourcePath!, target);
      if (!exists(fs, target))
        return this.guidance(
          `已请求${label}，但目标文件尚不能确认存在；请核对实际文件，不要视为完成。`,
        );
      const bytes = fileBytes(fs, target);
      if (
        (pdf && (!bytes.startsWith('%PDF-') || !/%%EOF\s*$/u.test(bytes.slice(-2048)))) ||
        (originalBytes !== undefined && bytes !== originalBytes) ||
        (!pdf &&
          originalBytes === undefined &&
          /\.(xlsx|xlsm|xlsb|xltx|xltm)$/iu.test(target) &&
          !bytes.startsWith('PK\u0003\u0004')) ||
        (!pdf &&
          originalBytes === undefined &&
          /\.xls$/iu.test(target) &&
          !bytes.startsWith('\u00d0\u00cf\u0011\u00e0\u00a1\u00b1\u001a\u00e1'))
      )
        return this.guidance(
          `已请求${label}，但文件内容核验未通过。可能留下未完成文件，请在 WPS 检查；助手不报告保存成功。`,
        );
      if (
        identity(source) !== sourceId ||
        this.host.documentId(this.host.application()) !== expected.documentId ||
        (sheetScope && !same(pdfSheetScope(this.host.application(), pdfWorkbook!), sheetScope)) ||
        (originalText !== undefined &&
          (sourceDocument?.Saved !== true ||
            sourceDocument.Content?.Text !== originalText ||
            fileBytes(fs, sourcePath!) !== originalBytes))
      )
        return this.guidance(
          `输出文件已出现，但原稿状态发生变化，请核对原稿和目标文件；助手不报告完成。`,
        );
      const actions: NonNullable<ChangeReceipt['actions']> = [];
      if (pdf && scope !== 'document') {
        try {
          assertSameSelection(expected, await this.host.readSelection(readOptions(expected)));
          const currentApp = this.host.application();
          if (this.host.kind === 'sheet')
            assertSinglePdfSheet(
              currentApp,
              pdfWorkbook!,
              sheetOf(currentApp, this.fixedRange(currentApp, expected)),
            );
          else {
            assertPdfTextBody(currentApp, writerSelection, expected, sourceId);
            assertPdfTextBody(currentApp, currentApp.Selection, expected, sourceId);
          }
        } catch {
          return this.guidance(
            `输出文件已出现，但选区、文字区域、工作表分组或内容在导出期间发生变化或无法核对。请核对“${target}”及原稿，不要直接重复导出；助手不报告完成。`,
          );
        }
      }
      // Retain at most one verified output's bytes, rather than accumulating large file buffers.
      this.results.clear();
      if (typeof app.OAAssist?.ShellExecute === 'function') {
        const open = `wps-result:${crypto.randomUUID()}`;
        this.results.set(open, { path: target, file: bytes });
        actions.push({ id: open, label: '打开结果' });
        if (directoryExists(fs, pathParent(target))) {
          const folder = `wps-result:${crypto.randomUUID()}`;
          this.results.set(folder, { path: pathParent(target) });
          actions.push({ id: folder, label: '打开文件夹' });
        }
      }
      return {
        token: `wps-task:${crypto.randomUUID()}`,
        canUndo: false,
        outcome: 'completed',
        location: target,
        message: `已${label}到“${target}”，并按字节读回核验。原稿名称与路径未改变。${sheetScope ? '已导出整个工作簿（所有工作表，含图表表），忽略各表打印区域。' : pdf ? `已导出${scopeLabel}。` : ''}${pdf ? '请打开 PDF 核对版式与分页。' : ''}`,
        ...(actions.length ? { actions } : {}),
      };
    } catch (error) {
      if (error instanceof HelperError && ['STALE_SELECTION', 'INVALID_PLAN'].includes(error.code))
        throw error;
      const reason =
        error instanceof HelperError ? error.message : 'WPS 原生接口没有完成可靠返回。';
      return this.guidance(
        attempted
          ? `已请求${label}，但未通过可靠核验：${reason}目标位置可能已生成文件，请人工检查；助手不报告完成。`
          : `${reason}${hint}`,
      );
    }
  }

  async openResult(actionId: string): Promise<void> {
    const result = this.results.get(actionId);
    if (!result) throw new HelperError('EXPIRED_TASK', '结果操作已失效，请重新保存并核验。');
    const app = this.host.application();
    const fs = app.FileSystem;
    if (
      !fs ||
      typeof app.OAAssist?.ShellExecute !== 'function' ||
      !absolutePath(result.path) ||
      !exists(fs, result.path)
    )
      throw unavailable('结果路径当前不可用，请从 WPS 原生界面打开。');
    if (result.file !== undefined && fileBytes(fs, result.path) !== result.file)
      throw new HelperError('STALE_SELECTION', '输出文件在保存后被修改，请自行核对。');
    if (result.file === undefined && !directoryExists(fs, result.path))
      throw unavailable('结果文件夹当前无法确认为目录，未调用打开操作。');
    // Only an adapter-owned path reaches this API; never execute caller-supplied paths or commands.
    await app.OAAssist.ShellExecute(result.path, '');
  }

  private notice(
    app: WpsApplication,
    plan: Extract<ExtendedActionPlan, { type: 'document.notice' }>,
  ): ChangeReceipt {
    if (this.host.kind !== 'text') throw unavailable('请在 WPS 文字中新建通知。');
    const content = noticeText(plan.fields).replace(/\r\n?/gu, '\n');
    const source = app.ActiveDocument;
    const documents = app.Documents;
    if (
      !source ||
      !source.Content ||
      typeof source.Content.Text !== 'string' ||
      typeof documents?.Add !== 'function'
    )
      return this.guidance(
        `当前版本无法核对新建文档接口。请在 WPS 文字新建空白文档后复制以下完整通知，原件未修改。\n\n${content}`,
      );
    const original = source.Content.Text;
    const sourceId = identity(source);
    const text = content.replace(/\n/gu, '\r');
    try {
      const target = documents.Add();
      if (
        !target ||
        target === source ||
        identity(target) === sourceId ||
        !target.Content ||
        !target.Content.Document ||
        identity(target.Content.Document) !== identity(target) ||
        typeof target.Content.Text !== 'string' ||
        !/^[\r\n]*$/u.test(target.Content.Text) ||
        target.ReadOnly !== false ||
        target.ProtectionType !== -1 ||
        !app.ActiveDocument ||
        identity(app.ActiveDocument) !== identity(target) ||
        identity(source) !== sourceId ||
        source.Content.Text !== original
      )
        throw unavailable();
      target.Content.Text = text;
      const normalized = target.Content.Text.replace(/\r\n?/gu, '\n').replace(/\n$/u, '');
      if (
        normalized !== text.replace(/\r/gu, '\n') ||
        identity(source) !== sourceId ||
        source.Content.Text !== original
      )
        throw unavailable();
      return {
        token: `wps-task:${crypto.randomUUID()}`,
        canUndo: false,
        outcome: 'completed',
        message: `已在新的未保存文档“${target.Name}”中生成通知。请检查后另存；原文档未修改。`,
      };
    } catch {
      throw new HelperError(
        'INVALID_PLAN',
        '新通知文档生成未通过核对。请查看可能已创建的未保存文档；助手没有保存文件或自动关闭文档。',
      );
    }
  }

  async locate(expected: SelectionSnapshot, row: number, column: number): Promise<void> {
    if (
      this.host.kind !== 'sheet' ||
      expected.kind !== 'sheet' ||
      !Number.isInteger(row) ||
      !Number.isInteger(column) ||
      row < 0 ||
      column < 0 ||
      !expected.values?.[row] ||
      column >= expected.values[row]!.length
    )
      throw new HelperError('INVALID_PLAN', '定位坐标超出原检查范围。');
    const app = this.host.application();
    if (this.host.documentId(app) !== expected.documentId) throw stale();
    const sheet = app.ActiveSheet;
    // Recover by the original sheet name/address; do not depend on the current selection.
    if (
      !sheet ||
      typeof sheet.Range !== 'function' ||
      !expected.targetId.startsWith(`${sheet.Name}!`)
    )
      throw stale();
    const address = expected.targetId.slice(sheet.Name.length + 1);
    const range = sheet.Range(address);
    const [rows, columns] = dimensions(range);
    if (
      addressOf(range) !== address ||
      sheetOf(app, range).Name !== sheet.Name ||
      !same(readWpsValues(range, rows, columns), expected.values)
    )
      throw stale();
    if (expected.cellTypes) {
      const details = readWpsSheetDetails(range, app, expected.values);
      if (!same(details.cellTypes, expected.cellTypes) || details.rowScope !== expected.rowScope)
        throw stale();
    }
    if (expected.sheetLayout && !same(readWpsSheetLayout(app, range), expected.sheetLayout))
      throw stale();
    if (expected.printSettings && !same(readWpsPrintSettings(app, range), expected.printSettings))
      throw stale();
    if (
      !Number.isInteger(range.Row) ||
      !Number.isInteger(range.Column) ||
      !range.Row ||
      !range.Column
    )
      throw unavailable('无法读取原检查范围的行列位置，未进行定位。');
    const targetRow = range.Row + row;
    const targetColumn = range.Column + column;
    const target = cellAt(range, row + 1, column + 1);
    if (typeof target.Select !== 'function')
      throw unavailable('当前 WPS 版本没有可核对的单元格定位接口。');
    target.Select();
    const currentApp = this.host.application();
    const selected = currentApp.Selection;
    if (
      this.host.documentId(currentApp) !== expected.documentId ||
      currentApp.ActiveSheet?.Name !== sheet.Name ||
      selected?.Row !== targetRow ||
      selected.Column !== targetColumn ||
      selected.Rows?.Count !== 1 ||
      selected.Columns?.Count !== 1 ||
      selected.Areas?.Count !== 1
    )
      throw unavailable('WPS 未返回预期的定位结果，请在原检查范围中手动查看该单元格。');
  }
}
