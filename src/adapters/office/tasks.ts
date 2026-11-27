import JSZip from 'jszip';
import type {
  Capability,
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
import { assertSameSelection } from '../../core/selection';

type Identity = Pick<SelectionSnapshot, 'documentId' | 'documentName'>;
type SheetPlan = Extract<ExtendedActionPlan, { type: `sheet.${string}` }>;
interface Origin {
  snapshot: SelectionSnapshot;
  sheetId: string;
  address: string;
  range?: Excel.Range;
  zoom?: Excel.PageLayoutZoomOptions;
}
interface Details {
  sheetLayout?: SheetLayoutSnapshot;
  printSettings?: PrintSettingsSnapshot;
  zoom?: Excel.PageLayoutZoomOptions;
}
interface SheetChange {
  token: string;
  origin: Origin;
  plan: SheetPlan;
  before: SelectionSnapshot;
  after: SelectionSnapshot;
  beforeZoom?: Excel.PageLayoutZoomOptions;
  afterZoom?: Excel.PageLayoutZoomOptions;
}
interface ExportResult {
  blob: Blob;
  name: string;
  pdf: boolean;
  handle?: SaveHandle;
  url?: string;
}

/** Requirements are tested independently; an older host still offers honest native guidance. */
function supported(set: string, version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported(set, version);
  } catch {
    return false;
  }
}

export function officeTaskCapabilities(kind: DocumentKind): Capability[] {
  if (typeof Office === 'undefined' || !Office.context.requirements) return [];
  const capabilities: Capability[] = ['document.saveCopy', 'document.exportPdf'];
  if (kind === 'text') {
    capabilities.push('document.notice');
  } else {
    capabilities.push('sheet.print', 'selection.locate');
    // Inspecting merged areas safely requires 1.13, even though the setters are older.
    if (supported('ExcelApi', '1.13'))
      capabilities.push('sheet.autofit', 'sheet.wrap', 'sheet.header', 'history.undo');
  }
  return capabilities;
}

function stale(message = '原文件、范围或数据已改变，请重新预览后再操作。'): never {
  throw new HelperError('STALE_SELECTION', message);
}
function unavailable(message: string): never {
  throw new HelperError('UNAVAILABLE', message);
}
function equal(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}
function key(snapshot: SelectionSnapshot): string {
  // No native handle or caller-owned mutable object is used as the authority.
  return JSON.stringify(snapshot);
}
function result(message: string, outcome: ChangeReceipt['outcome'] = 'guidance'): ChangeReceipt {
  return { token: crypto.randomUUID(), message, canUndo: false, outcome };
}

/** Native range/identity ownership stays here, including reports whose selection later changes. */
export class OfficeTasks {
  private readonly origins = new Map<string, Origin>();
  private change?: SheetChange;
  // Opaque action IDs own verified bytes and optional native handles; no path is invented.
  private readonly exportedResults = new Map<string, ExportResult>();
  private disposed = false;

  constructor(
    private readonly kind: DocumentKind,
    private readonly identity: () => Identity,
    private readonly verifyText: (expected: SelectionSnapshot) => Promise<void>,
  ) {}

  async capture(
    context: Excel.RequestContext,
    range: Excel.Range,
    sheetId: string,
    snapshot: SelectionSnapshot,
    options: SelectionReadOptions,
  ): Promise<SelectionSnapshot> {
    const details = await this.details(context, range, options);
    const captured = { ...snapshot, ...this.publicDetails(details) };
    const id = key(captured);
    if (!this.origins.has(id)) {
      // Existing read-only doubles/old hosts need no additional tracking API.
      const tracked = typeof range.track === 'function' ? range.track() : undefined;
      if (tracked) await context.sync();
      this.origins.set(id, {
        snapshot: structuredClone(captured),
        sheetId,
        address: snapshot.targetLabel,
        range: tracked,
        zoom: details.zoom,
      });
      // Bound report handles; undo owns its handle independently until consumed/disposed.
      if (this.origins.size > 16) {
        const oldest = this.origins.keys().next().value!;
        const origin = this.origins.get(oldest)!;
        this.origins.delete(oldest);
        if (origin !== this.change?.origin) this.release(origin);
      }
    }
    return captured;
  }

  async apply(plan: ExtendedActionPlan, expected: SelectionSnapshot): Promise<ChangeReceipt> {
    this.active();
    validateExtendedPlan(plan);
    if (expected.kind !== this.kind) stale();
    await this.verify(expected, true);
    if (plan.type === 'document.notice') return this.notice(plan, expected);
    if (plan.type === 'document.saveCopy' || plan.type === 'document.exportPdf')
      return this.exportFile(plan, expected);
    return this.applySheet(plan, expected);
  }

  owns(receipt: ChangeReceipt): boolean {
    return receipt.token.startsWith('office-task:');
  }

  async openResult(actionId: string): Promise<void> {
    this.active();
    const exported = this.exportedResults.get(actionId);
    if (!exported) throw new HelperError('EXPIRED_TASK', '这个导出结果已过期，请重新导出。');
    try {
      this.openBlob(exported, exported.pdf);
    } catch {
      unavailable(
        '浏览器未能打开结果。请在浏览器下载列表或原先选择的位置打开文件核对；加载项不能安全定位绝对文件夹。',
      );
    }
  }

  async undo(receipt: ChangeReceipt): Promise<void> {
    this.active();
    const change = this.change;
    if (!change || change.token !== receipt.token)
      throw new HelperError('EXPIRED_TASK', '这次修改已无法从助手中恢复。');
    await this.run(change.origin, async (context, range) => {
      const current = await this.state(context, range, change.origin, change.after);
      this.checkData(change.after, current.snapshot);
      this.guard(current.guard);
      // Compare only the properties this operation wrote. Other user formatting survives.
      if (
        !equal(
          this.property(current.snapshot, change.plan),
          this.property(change.after, change.plan),
        )
      )
        stale('本次修改的属性已有后续编辑，未覆盖；请使用原生撤销或重新预览。');
      if (change.plan.type === 'sheet.print' && !equal(current.zoom, change.afterZoom)) stale();
      try {
        await this.write(context, range, change.plan, change.before, change.beforeZoom);
        const restored = await this.state(context, range, change.origin, change.before);
        this.checkData(change.before, restored.snapshot);
        if (
          !equal(
            this.property(restored.snapshot, change.plan),
            this.property(change.before, change.plan),
          ) ||
          (change.plan.type === 'sheet.print' && !equal(restored.zoom, change.beforeZoom))
        )
          throw new Error('Restoration read-back differs');
      } catch {
        this.clearChange();
        throw new HelperError(
          'INVALID_PLAN',
          '恢复可能只完成了一部分，请检查原范围并使用 Excel 的原生撤销。',
        );
      }
    });
    this.clearChange();
  }

  async locate(snapshot: SelectionSnapshot, row: number, column: number): Promise<void> {
    this.active();
    const origin = this.origin(snapshot);
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(column) ||
      row < 0 ||
      column < 0 ||
      row >= (snapshot.values?.length ?? 0) ||
      column >= (snapshot.values?.[row]?.length ?? 0)
    )
      throw new HelperError('INVALID_PLAN', '定位位置超出了原始检查范围。');
    await this.run(origin, async (context, range) => {
      const current = await this.state(context, range, origin, snapshot, false);
      this.checkData(snapshot, current.snapshot);
      this.assertDocument(snapshot);
      range.worksheet.activate();
      range.getCell(row, column).select();
      await context.sync();
    });
  }

  dispose(): void {
    this.disposed = true;
    const origins = new Set(this.origins.values());
    if (this.change) origins.add(this.change.origin);
    for (const origin of origins) this.release(origin);
    this.origins.clear();
    this.change = undefined;
    for (const exported of this.exportedResults.values()) {
      if (exported.url) URL.revokeObjectURL(exported.url);
    }
    this.exportedResults.clear();
  }

  private async verify(expected: SelectionSnapshot, selection: boolean): Promise<void> {
    this.assertDocument(expected);
    if (this.kind === 'text') return this.verifyText(expected);
    const origin = this.origin(expected);
    await this.run(origin, async (context, range) => {
      const current = await this.state(context, range, origin, expected, false);
      assertSameSelection(expected, current.snapshot);
      if (selection) {
        const selected = context.workbook.getSelectedRange();
        const selectedSheet = selected.worksheet;
        selected.load('address');
        selectedSheet.load('id');
        await context.sync();
        if (selected.address !== origin.address || selectedSheet.id !== origin.sheetId) stale();
      }
      this.assertDocument(expected);
    });
  }

  private origin(snapshot: SelectionSnapshot): Origin {
    this.assertDocument(snapshot);
    const origin = this.origins.get(key(snapshot));
    if (!origin) stale('原始检查范围已失效，请重新检查后再操作。');
    return origin;
  }

  private async run<T>(
    origin: Origin,
    action: (ctx: Excel.RequestContext, range: Excel.Range) => Promise<T>,
  ): Promise<T> {
    this.assertDocument(origin.snapshot);
    const run = async (context: Excel.RequestContext) => {
      const range =
        origin.range ??
        context.workbook.worksheets
          .getItem(origin.sheetId)
          .getRange(origin.address.slice(origin.address.lastIndexOf('!') + 1));
      return action(context, range);
    };
    return origin.range ? Excel.run(origin.range, run) : Excel.run(run);
  }

  private async state(
    context: Excel.RequestContext,
    range: Excel.Range,
    origin: Origin,
    expected: SelectionSnapshot,
    guards = true,
  ) {
    const sheet = range.worksheet;
    range.load(['address', 'rowCount', 'columnCount', 'values']);
    sheet.load('id');
    if (expected.cellTypes) range.load('valueTypes');
    await context.sync();
    this.assertDocument(expected);
    if (
      range.address !== origin.address ||
      sheet.id !== origin.sheetId ||
      range.rowCount !== expected.values?.length ||
      range.columnCount !== expected.values?.[0]?.length
    )
      stale();
    const details = await this.details(context, range, {
      includeSheetLayout: !!expected.sheetLayout,
      includePrintSettings: !!expected.printSettings,
    });
    let guard = details.sheetLayout;
    if (guards && !guard) guard = await this.safety(context, range);
    const types = expected.cellTypes
      ? range.valueTypes.map((row) =>
          row.map((type) => {
            switch (type) {
              case 'Empty':
                return 'empty';
              case 'String':
                return 'text';
              case 'Integer':
              case 'Double':
                return 'number';
              case 'Boolean':
                return 'boolean';
              case 'Error':
                return 'error';
              default:
                return 'unknown';
            }
          }),
        )
      : undefined;
    const snapshot: SelectionSnapshot = {
      ...expected,
      ...this.identity(),
      values: structuredClone(range.values),
      revision: JSON.stringify(range.values),
      ...this.publicDetails(details),
      ...(types ? { cellTypes: types } : {}),
    };
    return { snapshot, zoom: details.zoom, guard };
  }

  private checkData(expected: SelectionSnapshot, actual: SelectionSnapshot): void {
    if (
      !equal(
        [expected.documentId, expected.targetId, expected.values, expected.cellTypes],
        [actual.documentId, actual.targetId, actual.values, actual.cellTypes],
      )
    )
      stale();
  }

  private publicDetails(details: Details) {
    return {
      ...(details.sheetLayout ? { sheetLayout: details.sheetLayout } : {}),
      ...(details.printSettings ? { printSettings: details.printSettings } : {}),
    };
  }

  private async safety(
    context: Excel.RequestContext,
    range: Excel.Range,
  ): Promise<SheetLayoutSnapshot> {
    if (!supported('ExcelApi', '1.13'))
      unavailable('当前版本不能可靠检查合并单元格，需 ExcelApi 1.13；请在 Excel 原生界面操作。');
    const sheet = range.worksheet;
    const protection = sheet.protection;
    const merged = range.getMergedAreasOrNullObject();
    merged.load('isNullObject');
    protection.load('protected');
    sheet.load('name');
    range.load(['rowIndex', 'columnIndex']);
    await context.sync();
    if (
      typeof merged.isNullObject !== 'boolean' ||
      typeof protection.protected !== 'boolean' ||
      typeof sheet.name !== 'string' ||
      !Number.isInteger(range.rowIndex) ||
      !Number.isInteger(range.columnIndex) ||
      range.rowIndex < 0 ||
      range.columnIndex < 0
    )
      unavailable('宿主未提供可靠的合并、保护或范围位置信息，未作修改。');
    return {
      merged: !merged.isNullObject,
      protected: protection.protected,
      firstRow: range.rowIndex,
      firstColumn: range.columnIndex,
      sheetName: sheet.name,
      columnWidths: [],
      widthUnit: 'points',
      wrapText: [],
      headerBold: [],
    };
  }

  private async details(
    context: Excel.RequestContext,
    range: Excel.Range,
    options: SelectionReadOptions,
  ): Promise<Details> {
    const details: Details = {};
    if (options.includeSheetLayout) {
      const layout = await this.safety(context, range);
      // Per-column/per-cell reads retain mixed original formatting in native point units.
      const formats: Excel.RangeFormat[] = [];
      for (let column = 0; column < range.columnCount; column++) {
        const format = range.getColumn(column).format;
        format.load('columnWidth');
        formats.push(format);
        if (column % 50 === 49) await context.sync();
      }
      await context.sync();
      layout.columnWidths = formats.map((format) => format.columnWidth);
      const cells: Excel.RangeFormat[][] = [];
      const headerFonts: Excel.RangeFont[] = [];
      let pending = 0;
      for (let row = 0; row < range.rowCount; row++) {
        cells[row] = [];
        for (let column = 0; column < range.columnCount; column++) {
          const format = range.getCell(row, column).format;
          format.load('wrapText');
          if (row === 0) {
            const font = format.font;
            font.load('bold');
            headerFonts.push(font);
          }
          cells[row]!.push(format);
          if (++pending % 50 === 0) await context.sync();
        }
      }
      await context.sync();
      layout.wrapText = cells.map((row) => row.map((cell) => cell.wrapText));
      layout.headerBold = headerFonts.map((font) => font.bold);
      if (
        layout.columnWidths.some((width) => !Number.isFinite(width) || width < 0) ||
        layout.wrapText.some((row) => row.some((wrap) => typeof wrap !== 'boolean')) ||
        layout.headerBold.some((bold) => typeof bold !== 'boolean')
      )
        unavailable('宿主未返回可安全恢复的原始列宽、换行或粗体信息，请使用 Excel 原生格式功能。');
      details.sheetLayout = layout;
    }
    if (options.includePrintSettings) {
      if (!supported('ExcelApi', '1.9')) {
        details.printSettings = {
          available: false,
          reason: '打印设置需要 ExcelApi 1.9，请按 Ctrl+P 在原生界面核对。',
          printArea: '',
          orientation: '',
          zoom: false,
          fitWide: 0,
          fitTall: 0,
        };
      } else {
        const layout = range.worksheet.pageLayout;
        layout.load(['orientation', 'zoom']);
        const area = layout.getPrintAreaOrNullObject();
        area.load(['isNullObject', 'address']);
        await context.sync();
        const zoom = structuredClone(layout.zoom);
        if (
          !zoom ||
          !['Portrait', 'Landscape'].includes(layout.orientation) ||
          !(
            (typeof zoom.scale === 'number' && zoom.scale >= 10 && zoom.scale <= 400) ||
            [zoom.horizontalFitToPages, zoom.verticalFitToPages].some(
              (value) => typeof value === 'number' && Number.isInteger(value) && value >= 0,
            )
          )
        )
          unavailable(
            '无法读取可安全恢复的打印缩放原值，请按 Ctrl+P 在 Excel 原生界面设置和核对。',
          );
        details.zoom = zoom;
        details.printSettings = {
          available: true,
          printArea: area.isNullObject ? '' : area.address,
          orientation: layout.orientation,
          zoom: typeof zoom.scale === 'number' ? zoom.scale : false,
          fitWide: zoom.horizontalFitToPages ?? 0,
          fitTall: zoom.verticalFitToPages ?? 0,
        };
      }
    }
    return details;
  }

  private guard(layout?: SheetLayoutSnapshot): void {
    if (!layout) unavailable('无法可靠检查原始范围的合并及保护状态，请在 Excel 中操作。');
    if (layout.merged || layout.protected)
      unavailable('范围含合并单元格或工作表受保护，未作修改；请先在 Excel 中处理合并或保护。');
  }

  private property(snapshot: SelectionSnapshot, plan: SheetPlan): unknown {
    switch (plan.type) {
      case 'sheet.autofit':
        return snapshot.sheetLayout?.columnWidths;
      case 'sheet.wrap':
        return snapshot.sheetLayout?.wrapText;
      case 'sheet.header':
        return snapshot.sheetLayout?.headerBold;
      case 'sheet.print':
        return snapshot.printSettings;
    }
  }

  private async applySheet(plan: SheetPlan, expected: SelectionSnapshot): Promise<ChangeReceipt> {
    if (this.kind !== 'sheet') unavailable('请在 Excel 中选中表格范围。');
    if (
      plan.type === 'sheet.print' &&
      (!supported('ExcelApi', '1.9') ||
        !supported('ExcelApi', '1.13') ||
        !expected.printSettings?.available)
    )
      return result(
        '当前宿主无法安全设置打印。请返回原始选区，依次选择「页面布局 → 打印区域 → 设置打印区域」，方向选「横向」，缩放设「宽度 1 页、高度 自动」；按 Ctrl+P 核对实际页数、范围和纸张，再由你确认打印。助手没有提交打印。',
      );
    if (plan.type === 'sheet.print' ? !expected.printSettings?.available : !expected.sheetLayout)
      unavailable('预览缺少原始布局或打印设置，请重新预览。');
    const origin = this.origin(expected);
    return this.run(origin, async (context, range) => {
      const before = await this.state(context, range, origin, expected);
      assertSameSelection(expected, before.snapshot);
      this.guard(before.guard);
      this.clearChange();
      try {
        await this.write(context, range, plan);
        const after = await this.state(context, range, origin, expected);
        this.checkData(expected, after.snapshot);
        const desired =
          plan.type === 'sheet.wrap'
            ? after.snapshot.sheetLayout?.wrapText.every((row) => row.every(Boolean))
            : plan.type === 'sheet.header'
              ? after.snapshot.sheetLayout?.headerBold.every(Boolean)
              : plan.type === 'sheet.print'
                ? after.snapshot.printSettings?.orientation === 'Landscape' &&
                  after.snapshot.printSettings.zoom === false &&
                  after.snapshot.printSettings.fitWide === 1 &&
                  after.snapshot.printSettings.fitTall === 0 &&
                  after.snapshot.printSettings.printArea.replace(/\$/gu, '') ===
                    range.address.replace(/\$/gu, '')
                : after.snapshot.sheetLayout?.columnWidths.every(
                    (width) => Number.isFinite(width) && width > 0,
                  );
        if (!desired) throw new Error('Native read-back differs');
        const token = `office-task:${crypto.randomUUID()}`;
        this.change = {
          token,
          origin,
          plan,
          before: before.snapshot,
          after: after.snapshot,
          beforeZoom: before.zoom,
          afterZoom: after.zoom,
        };
        return {
          token,
          canUndo: true,
          outcome: 'completed',
          message:
            plan.type === 'sheet.print'
              ? '已设置原选区打印区域、横向、宽 1 页且高度自动。实际页数需按 Ctrl+P 在 Excel 原生打印预览核对，并由你确认打印；助手没有提交打印。'
              : plan.type === 'sheet.autofit'
                ? '已自动调整原选区涉及列的列宽（影响整列），可恢复原列宽。'
                : plan.type === 'sheet.wrap'
                  ? '已为原选区启用自动换行，可恢复各单元格原值。'
                  : '已将原选区首行加粗，可恢复各单元格原粗体状态。',
        };
      } catch {
        throw new HelperError(
          'INVALID_PLAN',
          'Excel 可能已作部分修改，但未能完成结果检查；请检查原范围并使用 Excel 的原生撤销。',
        );
      }
    });
  }

  private async write(
    context: Excel.RequestContext,
    range: Excel.Range,
    plan: SheetPlan,
    restore?: SelectionSnapshot,
    zoom?: Excel.PageLayoutZoomOptions,
  ): Promise<void> {
    if (plan.type === 'sheet.print') {
      range.worksheet.pageLayout.setPrintArea(restore ? restore.printSettings!.printArea : range);
      range.worksheet.pageLayout.orientation = (restore?.printSettings?.orientation ??
        'Landscape') as Excel.PageOrientation;
      range.worksheet.pageLayout.zoom = restore
        ? { ...zoom! }
        : { horizontalFitToPages: 1, verticalFitToPages: 0 };
    } else if (plan.type === 'sheet.autofit') {
      if (!restore) range.format.autofitColumns();
      else
        for (let column = 0; column < range.columnCount; column++) {
          range.getColumn(column).format.columnWidth = restore.sheetLayout!.columnWidths[column]!;
          if (column % 50 === 49) await context.sync();
        }
    } else if (plan.type === 'sheet.wrap') {
      if (!restore) range.format.wrapText = true;
      else {
        let pending = 0;
        for (let row = 0; row < range.rowCount; row++) {
          for (let column = 0; column < range.columnCount; column++) {
            range.getCell(row, column).format.wrapText =
              restore.sheetLayout!.wrapText[row]![column]!;
            if (++pending % 50 === 0) await context.sync();
          }
        }
      }
    } else if (plan.type === 'sheet.header') {
      if (!restore) range.getRow(0).format.font.bold = true;
      else
        for (let column = 0; column < range.columnCount; column++) {
          range.getCell(0, column).format.font.bold = restore.sheetLayout!.headerBold[column]!;
          if (column % 50 === 49) await context.sync();
        }
    }
    await context.sync();
  }

  private async notice(
    plan: Extract<ExtendedActionPlan, { type: 'document.notice' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    if (this.kind !== 'text') unavailable('请从 Word 生成通知。');
    const text = noticeText(plan.fields);
    if (!supported('WordApi', '1.3'))
      return result(
        `当前 Word 无法从加载项创建新文档。请按 Ctrl+N 新建文档，将以下通知粘贴进去，再用 F12 另存；原件保持不变。\n\n${text}`,
      );
    const xml = (value: string) =>
      value.replace(
        /[&<>"']/gu,
        (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[char]!,
      );
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    );
    zip.file(
      '_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    );
    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${text
        .replace(/\r\n?/gu, '\n')
        .split('\n')
        .map(
          (line, index) =>
            `<w:p>${index === 0 ? '<w:pPr><w:jc w:val="center"/></w:pPr>' : ''}<w:r>${index === 0 ? '<w:rPr><w:b/><w:sz w:val="36"/></w:rPr>' : ''}<w:t xml:space="preserve">${xml(line)}</w:t></w:r></w:p>`,
        )
        .join('')}<w:sectPr/></w:body></w:document>`,
    );
    const base64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
    try {
      await Word.run(async (context) => {
        this.assertDocument(expected);
        const created = context.application.createDocument(base64);
        await context.sync();
        // DocumentCreated.body is WordApiHiddenDocument 1.3, NOT WordApi 1.3.
        if (supported('WordApiHiddenDocument', '1.3')) {
          created.body.load('text');
          await context.sync();
          const normalize = (value: string) => value.replace(/\r\n?/gu, '\n').trim();
          if (normalize(created.body.text) !== normalize(text))
            throw new Error('Notice read-back differs');
        }
        created.open();
        await context.sync();
      });
      return result(
        '已在新的 Word 文档中打开通知，原件未改动。请在新文档核对内容并另存。',
        'completed',
      );
    } catch {
      return result(
        `未能确认新文档已打开，请先检查是否已有新窗口。需要手动生成时，按 Ctrl+N，新建后粘贴以下通知并另存；不要覆盖原件。\n\n${text}`,
      );
    }
  }

  private async exportFile(
    plan: Extract<ExtendedActionPlan, { type: 'document.saveCopy' | 'document.exportPdf' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    const pdf = plan.type === 'document.exportPdf';
    const scope = plan.type === 'document.exportPdf' ? (plan.scope ?? 'document') : 'document';
    const nativeScope =
      scope === 'selection'
        ? '选定内容（当前选区）'
        : scope === 'activeSheet'
          ? '活动工作表'
          : '整个工作簿（所有工作表）';
    const guidance = () =>
      result(
        `${pdf ? '尚未验证 PDF 已保存' : '尚未验证副本已保存'}。请回到${this.kind === 'text' ? ' Word' : ' Excel'}，选择「文件 → ${pdf ? '导出 → 创建 PDF/XPS（或另存为 → PDF）' : '另存为/保存副本 → 浏览'}」，${plan.folder ? `目标文件夹选择「${plan.folder}」，` : '选择目标文件夹，'}文件名填写「${plan.fileName}」${pdf ? (this.kind === 'sheet' ? `，在「选项 → ${nativeScope}」确认范围、核对页面后发布` : scope === 'selection' ? '，在「选项 → 选定内容」确认仅导出当前选区，检查预览后发布' : '，检查导出页面后发布') : '，保留原文件类型并保存到不同于原件的位置'}。完成后在该位置重新打开文件核对；原件不要覆盖。`,
      );
    // GetFileOptions has no PDF scope option, and the generic "entire document"
    // description does not specify Excel worksheet/print-area semantics. Do not
    // create or download bytes we cannot verify against the whole-workbook promise.
    if (pdf && (this.kind === 'sheet' || scope !== 'document')) {
      const receipt = guidance();
      receipt.message = `无法可靠确认自动导出的 PDF ${scope === 'document' ? '包含整个工作簿（所有工作表）' : `仅包含${nativeScope}`}，助手未自动导出或保存。${receipt.message}`;
      return receipt;
    }
    const picker = (
      globalThis as typeof globalThis & {
        showSaveFilePicker?: (options: unknown) => Promise<SaveHandle>;
      }
    ).showSaveFilePicker;
    if (plan.folder?.trim() || typeof Office.context.document.getFileAsync !== 'function')
      return guidance();
    const sourceNames = new Set([expected.documentName.toLowerCase()]);
    try {
      const sourceUrl = new URL(Office.context.document.url);
      if (['http:', 'https:', 'file:'].includes(sourceUrl.protocol))
        sourceNames.add(
          decodeURIComponent(sourceUrl.pathname.split('/').pop() || '').toLowerCase(),
        );
    } catch {
      /* A native unsaved document or Windows path need not be a URL. */
    }
    const extension = pdf
      ? '.pdf'
      : this.kind === 'text'
        ? '.docx'
        : [...sourceNames].some((name) => name.endsWith('.xlsm'))
          ? '.xlsm'
          : '.xlsx';
    if (!plan.fileName.toLowerCase().endsWith(extension)) return guidance();
    if (
      !pdf &&
      (sourceNames.has(plan.fileName.toLowerCase()) ||
        [...sourceNames].some((name) =>
          /\.(?:docm|dotm|dotx|xlsb|xls|xltm|xltx|xlam)$/iu.test(name),
        ))
    )
      return guidance();
    if (!picker) {
      // A download is only attempted after a successful native export, never after unsupported.
      try {
        if (
          typeof document === 'undefined' ||
          !document.body ||
          typeof URL.createObjectURL !== 'function'
        )
          return guidance();
        const bytes = await getOfficeFile(pdf);
        await this.verify(expected, false);
        this.active();
        const blob = new Blob([new Uint8Array(bytes)], {
          type: pdf ? 'application/pdf' : 'application/octet-stream',
        });
        const exported = { blob, name: plan.fileName, pdf };
        this.openBlob(exported, false);
        return {
          ...result(
            `已将${pdf ? ' PDF' : '副本'}「${plan.fileName}」交给浏览器下载，尚未确认文件已保存。请在下载列表选择保存位置并打开核对；浏览器不向助手披露绝对路径，也不能由加载项定位绝对文件夹。`,
            'download',
          ),
          location: plan.fileName,
          actions: this.retainExport(exported),
        };
      } catch {
        return guidance();
      }
    }
    let handle: SaveHandle;
    try {
      handle = await picker({
        suggestedName: plan.fileName,
        excludeAcceptAllOption: true,
        types: [
          {
            description: pdf ? 'PDF' : 'Office 副本',
            accept: { [pdf ? 'application/pdf' : 'application/octet-stream']: [extension] },
          },
        ],
      });
    } catch (error) {
      if (error instanceof Error && error.name === 'AbortError')
        return result('已取消选择保存位置，未保存文件。', 'cancelled');
      return guidance();
    }
    let stream: SaveStream | undefined;
    let wrote = false;
    try {
      const bytes = await getOfficeFile(pdf);
      await this.verify(expected, false);
      this.active();
      if (!handle.name || handle.name.toLowerCase() !== plan.fileName.toLowerCase())
        return guidance();
      // A picker may return the source file. Reject known matching names conservatively.
      if (sourceNames.has(handle.name.toLowerCase())) return guidance();
      stream = await handle.createWritable();
      wrote = true;
      await stream.write(bytes);
      await stream.close();
      stream = undefined;
      const saved = new Uint8Array(await (await handle.getFile()).arrayBuffer());
      if (saved.length !== bytes.length || saved.some((byte, index) => byte !== bytes[index]))
        throw new Error('Saved bytes differ');
      const receipt = {
        ...result(
          `已保存并回读核对${pdf ? ' PDF' : '副本'}「${handle.name}」的文件字节。浏览器不向助手披露绝对路径，显示的位置仅为文件名；加载项不能安全定位绝对文件夹。`,
          'completed',
        ),
        location: handle.name,
        actions: this.retainExport({
          blob: new Blob([new Uint8Array(bytes)], {
            type: pdf ? 'application/pdf' : 'application/octet-stream',
          }),
          name: handle.name,
          pdf,
          handle,
        }),
      };
      return receipt;
    } catch {
      if (stream) {
        try {
          await stream.abort();
        } catch {
          /* A failed/closed native writer may reject abort too. */
        }
      }
      const fallback = guidance();
      if (wrote) fallback.message = `目标位置可能已有部分文件，请先检查并处理。${fallback.message}`;
      return fallback;
    }
  }

  private retainExport(exported: ExportResult): NonNullable<ChangeReceipt['actions']> {
    const id = `office-result:${crypto.randomUUID()}`;
    this.exportedResults.set(id, exported);
    // Retain at most four results and 128 MiB, including their optional native handles.
    let size = [...this.exportedResults.values()].reduce(
      (total, item) => total + item.blob.size,
      0,
    );
    while (this.exportedResults.size > 4 || size > 128 * 1024 * 1024) {
      const oldest = this.exportedResults.keys().next().value!;
      const removed = this.exportedResults.get(oldest)!;
      this.exportedResults.delete(oldest);
      if (removed.url) URL.revokeObjectURL(removed.url);
      size -= removed.blob.size;
    }
    return [
      { id, label: exported.pdf ? '打开导出 PDF 预览（请核对）' : '重新下载副本（由浏览器保存）' },
    ];
  }

  private openBlob(exported: ExportResult, preview: boolean): void {
    if (typeof URL.createObjectURL !== 'function') throw new Error('Blob URLs unavailable');
    const url = (exported.url ??= URL.createObjectURL(exported.blob));
    try {
      if (preview && typeof window !== 'undefined' && typeof window.open === 'function') {
        // This requests a preview; it does not claim that a browser tab was opened or a file saved.
        window.open(url, '_blank', 'noopener,noreferrer');
        return;
      }
      if (typeof document === 'undefined' || !document.body)
        throw new Error('Browser UI unavailable');
      const anchor = document.createElement('a');
      try {
        anchor.href = url;
        if (preview) {
          anchor.target = '_blank';
          anchor.rel = 'noopener noreferrer';
        } else anchor.download = exported.name;
        anchor.hidden = true;
        document.body.appendChild(anchor);
        anchor.click();
      } finally {
        anchor.remove();
      }
    } catch (error) {
      URL.revokeObjectURL(url);
      exported.url = undefined;
      throw error;
    }
  }

  private assertDocument(expected: SelectionSnapshot): void {
    this.active();
    if (this.identity().documentId !== expected.documentId)
      stale('当前文件身份已改变，请回到原文件重新检查。');
  }
  private clearChange(): void {
    const previous = this.change;
    this.change = undefined;
    if (previous && ![...this.origins.values()].includes(previous.origin))
      this.release(previous.origin);
  }
  private active(): void {
    if (this.disposed) throw new HelperError('EXPIRED_TASK', '请重新打开助手。');
  }
  private release(origin: Origin): void {
    if (origin.range)
      void Excel.run(origin.range, async (context) => {
        origin.range!.untrack();
        await context.sync();
      }).catch(() => undefined);
  }
}

interface SaveStream {
  write(data: Uint8Array): Promise<void>;
  close(): Promise<void>;
  abort(): Promise<void>;
}
interface SaveHandle {
  name: string;
  createWritable(): Promise<SaveStream>;
  getFile(): Promise<{ arrayBuffer(): Promise<ArrayBuffer> }>;
}

function officeCallback<T>(
  invoke: (callback: (result: Office.AsyncResult<T>) => void) => void,
  late?: (value: T) => void,
): Promise<T> {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      settled = true;
      reject(new Error('Office file API timed out'));
    }, 15000);
    try {
      invoke((response) => {
        const succeeded = response.status === Office.AsyncResultStatus.Succeeded;
        if (settled) {
          if (succeeded) {
            try {
              late?.(response.value);
            } catch {
              /* Best-effort release after timeout. */
            }
          }
          return;
        }
        settled = true;
        clearTimeout(timer);
        if (succeeded) resolve(response.value);
        else reject(new Error(response.error?.message || 'Office file API failed'));
      });
    } catch (error) {
      settled = true;
      clearTimeout(timer);
      reject(error);
    }
  });
}

async function getOfficeFile(pdf: boolean): Promise<Uint8Array> {
  // 64 KiB is supported on iPad as well as desktop; always close the native File.
  const file = await officeCallback<Office.File>(
    (callback) =>
      Office.context.document.getFileAsync(
        pdf ? Office.FileType.Pdf : Office.FileType.Compressed,
        { sliceSize: 65536 },
        callback,
      ),
    (late) => late.closeAsync(),
  );
  try {
    if (
      !Number.isSafeInteger(file.size) ||
      file.size <= 0 ||
      file.size > 128 * 1024 * 1024 ||
      !Number.isSafeInteger(file.sliceCount) ||
      file.sliceCount <= 0 ||
      file.sliceCount > 2048
    )
      throw new Error('Invalid Office file size');
    const bytes = new Uint8Array(file.size);
    let offset = 0;
    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await officeCallback<Office.Slice>((callback) =>
        file.getSliceAsync(index, callback),
      );
      if (
        slice.index !== index ||
        !Array.isArray(slice.data) ||
        slice.data.length !== slice.size ||
        slice.data.some(
          (byte: unknown) =>
            typeof byte !== 'number' || !Number.isInteger(byte) || byte < 0 || byte > 255,
        ) ||
        offset + slice.size > bytes.length
      )
        throw new Error('Invalid Office slice');
      bytes.set(slice.data, offset);
      offset += slice.size;
    }
    if (
      offset !== bytes.length ||
      (pdf
        ? String.fromCharCode(...bytes.slice(0, 5)) !== '%PDF-'
        : bytes[0] !== 0x50 || bytes[1] !== 0x4b)
    )
      throw new Error('Unexpected exported file');
    return bytes;
  } finally {
    await officeCallback<void>((callback) => file.closeAsync(callback));
  }
}
