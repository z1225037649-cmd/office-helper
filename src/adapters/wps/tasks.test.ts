import { describe, expect, it, vi } from 'vitest';
import { createRuntime } from '../../app/runtime';
import type { CellValue, ExtendedActionPlan, NoticeFields } from '../../core/contracts';
import { noticeText } from '../../core/extended-actions';
import { WpsAdapter } from './adapter';
import type {
  WpsApplication,
  WpsDocument,
  WpsPageSetup,
  WpsRange,
  WpsSelection,
  WpsSheet,
  WpsSheetPdfSource,
  WpsSheetCell,
  WpsWorkbook,
  WpsWorksheet,
  WpsWriterPdfSelection,
} from './bridge';

function sheetFixture() {
  const state = {
    values: [
      ['项目', '金额'],
      ['甲', 12],
      ['乙', 7],
    ] as CellValue[][],
    widths: [10, 18],
    wraps: [
      [false, true],
      [true, false],
      [false, true],
    ],
    bold: [false, true],
    merged: false,
    writes: [] as string[],
    selected: [] as number[],
    ignoreWrap: false,
    failWrap: false,
  };
  const book: WpsWorkbook = { Name: '费用.xlsx', FullName: 'C:\\work\\费用.xlsx', ReadOnly: false };
  const page: WpsPageSetup = {
    PrintArea: '$A$1:$F$20',
    Orientation: 1,
    Zoom: 75,
    FitToPagesWide: 2,
    FitToPagesTall: 3,
  };
  const sheet: WpsWorksheet = {
    Name: '费用',
    Parent: book,
    ProtectContents: false,
    PageSetup: page,
  };
  function cell(r: number, c: number): WpsSheetCell {
    return {
      get Value2() {
        return state.values[r - 1]![c - 1]!;
      },
      get ColumnWidth() {
        return state.widths[c - 1]!;
      },
      set ColumnWidth(value) {
        state.writes.push(`width:${c}`);
        state.widths[c - 1] = value!;
      },
      get MergeCells() {
        return state.merged;
      },
      get WrapText() {
        return state.wraps[r - 1]![c - 1]!;
      },
      set WrapText(value) {
        state.writes.push(`wrap:${r}:${c}`);
        if (!state.ignoreWrap) state.wraps[r - 1]![c - 1] = Boolean(value);
        if (state.failWrap) throw new Error('partial write');
      },
      Font: {
        get Bold() {
          return state.bold[c - 1] ? -1 : 0;
        },
        set Bold(value: number | boolean | undefined) {
          state.writes.push(`bold:${r}:${c}`);
          state.bold[c - 1] = value === -1 || value === true;
        },
      },
      Select: vi.fn(() => {
        state.selected = [r, c];
        app.Selection = {
          Row: range.Row! + r - 1,
          Column: range.Column! + c - 1,
          Rows: { Count: 1 },
          Columns: { Count: 1 },
          Areas: { Count: 1 },
          Worksheet: sheet,
        };
      }),
    };
  }
  const range: WpsSelection = {
    Rows: { Count: 3 },
    Columns: {
      Count: 2,
      AutoFit: vi.fn(() => {
        state.writes.push('autofit');
        state.widths[0] = 20;
      }),
    },
    Areas: { Count: 1 },
    Row: 2,
    Column: 2,
    Address: '$B$2:$C$4',
    Worksheet: sheet,
    Cells: { Item: cell },
    get Value2() {
      return structuredClone(state.values);
    },
  };
  sheet.Range = vi.fn((address) => {
    if (address !== '$B$2:$C$4') throw new Error(`Unexpected range ${address}`);
    return range;
  });
  const app: WpsApplication = {
    ActiveWorkbook: book,
    ActiveSheet: sheet,
    Selection: range,
    ActiveWindow: { Hwnd: 9 },
    Enum: {
      xlLandscape: 2,
      xlPortrait: 1,
      xlTypePDF: 0,
      wdExportFormatPDF: 17,
      msoFileDialogSaveAs: 2,
    },
    WorksheetFunction: { IsError: () => false },
  };
  const host = new WpsAdapter('sheet', () => app);
  const preview = () => host.readSelection({ includeSheetLayout: true });
  const printPreview = () =>
    host.readSelection({ includeSheetLayout: true, includePrintSettings: true });
  return { state, app, book, sheet, range, page, host, preview, printPreview };
}

function files(app: WpsApplication) {
  const data = new Map<string, string>([['C:\\work\\费用.xlsx', 'PK\u0003\u0004source']]);
  const folders = new Set(['C:\\out', 'C:\\work']);
  const fs = {
    Exists: vi.fn((path: string) => data.has(path) || folders.has(path)),
    readAsBinaryString: vi.fn((path: string) => data.get(path)),
    copyFileSync: vi.fn((source: string, target: string) => {
      data.set(target, data.get(source)!);
    }),
    readdirSync: vi.fn((path: string): unknown => {
      if (!folders.has(path)) throw new Error('not a directory');
      return [];
    }),
  };
  app.FileSystem = fs;
  const open = vi.fn();
  app.OAAssist = { ShellExecute: open };
  return { data, folders, fs, open };
}

function pdfFixture() {
  const f = sheetFixture();
  const v = files(f.app);
  // Deliberately nonstandard enum values prove that export never guesses them.
  f.app.Enum!.xlSheetVisible = 47;
  f.app.Enum!.xlTypePDF = 91;
  f.sheet.Visible = 47;
  f.app.ActiveWindow!.SelectedSheets = { Count: 1, Item: () => f.sheet };
  const sheets: WpsSheet[] = [
    f.sheet,
    { Name: '附表', Parent: f.book, Visible: 47 },
    { Name: '图表页', Parent: f.book, Visible: 47 },
  ];
  const item = vi.fn((index: number) => sheets[index - 1]!);
  f.book.Sheets = {
    get Count() {
      return sheets.length;
    },
    Item: item,
  };
  const exportPdf = vi.fn((_type: number, path: string) => {
    v.data.set(path, '%PDF-1.7\nwhole workbook\n%%EOF\n');
  });
  f.book.ExportAsFixedFormat = exportPdf;
  const show = vi.fn(() => -1);
  f.app.FileDialog = vi.fn(() => ({
    InitialFileName: '',
    Title: '',
    ButtonName: '',
    Show: show,
    SelectedItems: { Count: 1, Item: () => 'C:\\out\\输出.pdf' },
  }));
  return { ...f, ...v, sheets, item, exportPdf, show };
}

const fields: NoticeFields = {
  title: '活动通知',
  time: '明天上午九点',
  location: '会议室',
  body: '请准时参加。\n请携带材料。',
  contact: '办公室',
};

function textFixture() {
  let sourceText = '原始正文\r';
  let sourceWrites = 0;
  const source: WpsDocument = {
    Name: '原稿.docx',
    FullName: 'C:\\work\\原稿.docx',
    ReadOnly: false,
    ProtectionType: -1,
    Saved: true,
  };
  source.Content = {
    Document: source,
    get Text() {
      return sourceText;
    },
    set Text(value) {
      sourceWrites++;
      sourceText = value;
    },
  };
  const target: WpsDocument = { Name: '文档2', ReadOnly: false, ProtectionType: -1 };
  let targetText = '\r';
  target.Content = {
    Document: target,
    get Text() {
      return targetText;
    },
    set Text(value) {
      targetText = value + '\r';
    },
  };
  const app: WpsApplication = {
    ActiveDocument: source,
    ActiveWindow: { Hwnd: 8 },
    Enum: { wdExportFormatPDF: 17, wdMainTextStory: 41 },
    Selection: { Start: 0, End: 4, Text: '原始正文', Font: { Size: 12, Bold: 0 } },
  };
  const range: WpsRange = {
    Text: '原始正文',
    Start: 0,
    End: 4,
    StoryType: 41,
    Font: app.Selection!.Font as WpsRange['Font'],
    Document: source,
    get Duplicate() {
      return this;
    },
  };
  app.Selection!.Range = range;
  const add = vi.fn(() => {
    app.ActiveDocument = target;
    return target;
  });
  app.Documents = { Add: add };
  const host = new WpsAdapter('text', () => app);
  return { app, source, target, add, host, sourceWrites: () => sourceWrites };
}

describe('WPS 扩展格式和逐属性安全恢复', () => {
  it('版式和打印状态只按需读取，原始混合格式逐列逐格保存', async () => {
    const { host, preview, printPreview } = sheetFixture();
    expect(await host.readSelection()).not.toHaveProperty('sheetLayout');
    expect(await host.readSelection()).not.toHaveProperty('printSettings');
    expect((await preview()).sheetLayout).toEqual({
      columnWidths: [10, 18],
      widthUnit: 'characters',
      wrapText: [
        [false, true],
        [true, false],
        [false, true],
      ],
      headerBold: [false, true],
      merged: false,
      protected: false,
      firstRow: 2,
      firstColumn: 2,
      sheetName: '费用',
    });
    expect((await printPreview()).printSettings).toMatchObject({
      available: true,
      zoom: 75,
      fitTall: 3,
    });
  });

  it.each(['sheet.autofit', 'sheet.wrap', 'sheet.header'] as const)(
    '%s 执行并恢复原属性，保留其他属性后续修改',
    async (type) => {
      const f = sheetFixture();
      const before = await f.preview();
      const receipt = await f.host.apply({ type }, before);
      expect(receipt).toMatchObject({ outcome: 'completed', canUndo: true });
      if (type === 'sheet.autofit') f.state.bold[0] = true;
      else f.state.widths[0] = 29;
      f.app.Selection = { ...f.range, Address: '$F$2:$G$4', Column: 6 };
      await f.host.undo(receipt);
      if (type === 'sheet.autofit') {
        expect(f.state.widths).toEqual([10, 18]);
        expect(f.state.bold).toEqual([true, true]);
      } else {
        expect(f.state.widths).toEqual([29, 18]);
        expect(type === 'sheet.wrap' ? f.state.wraps : f.state.bold).toEqual(
          type === 'sheet.wrap' ? before.sheetLayout!.wrapText : before.sheetLayout!.headerBold,
        );
      }
      expect(f.state.values).toEqual(before.values);
      await expect(f.host.undo(receipt)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    },
  );

  it('恢复不写入本次原本未改变的列、单元格或表头', async () => {
    for (const type of ['sheet.autofit', 'sheet.wrap', 'sheet.header'] as const) {
      const f = sheetFixture();
      const receipt = await f.host.apply({ type }, await f.preview());
      if (type === 'sheet.autofit') f.state.widths[1] = 31;
      else if (type === 'sheet.wrap') f.state.wraps[0]![1] = false;
      else f.state.bold[1] = false;
      f.state.writes = [];
      await f.host.undo(receipt);
      if (type === 'sheet.autofit') {
        expect(f.state.widths).toEqual([10, 31]);
        expect(f.state.writes).toEqual(['width:1']);
      } else if (type === 'sheet.wrap') {
        expect(f.state.wraps[0]![1]).toBe(false);
        expect(f.state.writes).not.toContain('wrap:1:2');
      } else {
        expect(f.state.bold).toEqual([false, false]);
        expect(f.state.writes).toEqual(['bold:1:1']);
      }
    }
  });

  it.each(['sheet.autofit', 'sheet.wrap', 'sheet.header'] as const)(
    '%s 恢复前发现目标属性冲突则零写入',
    async (type) => {
      const f = sheetFixture();
      const receipt = await f.host.apply({ type }, await f.preview());
      if (type === 'sheet.autofit') f.state.widths[0] = 25;
      else if (type === 'sheet.wrap') f.state.wraps[1]![1] = false;
      else f.state.bold[0] = false;
      f.state.writes = [];
      await expect(f.host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(f.state.writes).toEqual([]);
    },
  );

  it.each(['values', 'layout', 'position', 'document'] as const)(
    '预览后 %s 改变拒绝写入',
    async (what) => {
      const f = sheetFixture();
      const before = await f.preview();
      if (what === 'values') f.state.values[1]![1] = 999;
      if (what === 'layout') f.state.widths[1] = 23;
      if (what === 'position') f.app.Selection = { ...f.range, Address: '$D$2:$E$4' };
      if (what === 'document') f.book.FullName = 'C:\\work\\other.xlsx';
      await expect(f.host.apply({ type: 'sheet.wrap' }, before)).rejects.toMatchObject({
        code: 'STALE_SELECTION',
      });
      expect(f.state.writes).toEqual([]);
    },
  );

  it.each([
    'merged',
    'protected',
    'readonly',
    'areas',
    'missing-cells',
    'missing-protection',
    'missing-autofit',
  ] as const)('拒绝 %s 并保持零写入', async (what) => {
    const f = sheetFixture();
    if (what === 'merged') f.state.merged = true;
    if (what === 'protected') f.sheet.ProtectContents = true;
    if (what === 'readonly') f.book.ReadOnly = true;
    if (what === 'areas') f.range.Areas!.Count = 2;
    if (what === 'missing-cells') f.range.Cells = undefined;
    if (what === 'missing-protection') f.sheet.ProtectContents = undefined;
    if (what === 'missing-autofit') f.range.Columns!.AutoFit = undefined;
    await expect(
      (async () => f.host.apply({ type: 'sheet.autofit' }, await f.preview()))(),
    ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    expect(f.state.writes).toEqual([]);
  });

  it('固定目标必须与原始选区匹配', async () => {
    const f = sheetFixture();
    const before = await f.preview();
    f.sheet.Range = () => ({ ...f.range, Address: '$D$2:$E$4' });
    await expect(f.host.apply({ type: 'sheet.header' }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(f.state.writes).toEqual([]);
  });

  it.each(['partial', 'ignored'] as const)('原生 %s 写入不发成功凭证，旧凭证失效', async (mode) => {
    const f = sheetFixture();
    const old = await f.host.apply({ type: 'sheet.autofit' }, await f.preview());
    f.state.failWrap = mode === 'partial';
    f.state.ignoreWrap = mode === 'ignored';
    await expect(f.host.apply({ type: 'sheet.wrap' }, await f.preview())).rejects.toMatchObject({
      code: 'INVALID_PLAN',
    });
    await expect(f.host.undo(old)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
  });

  it('普通读取保留非有限数值；格式写入另行拒绝', async () => {
    const f = sheetFixture();
    f.state.values[1]![1] = NaN;
    expect((await f.host.readSelection({ includeCellTypes: true })).cellTypes?.[1]?.[1]).toBe(
      'unknown',
    );
    await expect(f.host.apply({ type: 'sheet.wrap' }, await f.preview())).rejects.toMatchObject({
      code: 'UNAVAILABLE',
    });
    expect(f.state.writes).toEqual([]);
  });
});

describe('WPS 扩展边界补充', () => {
  it('null与NaN不能因JSON序列化相同而通过固定范围核验', async () => {
    const f = sheetFixture();
    f.state.values[1]![1] = null;
    const before = await f.preview();
    const readRange = f.sheet.Range!;
    f.sheet.Range = (address) => {
      const result = readRange(address);
      f.state.values[1]![1] = NaN;
      return result;
    };
    await expect(f.host.apply({ type: 'sheet.wrap' }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(f.state.writes).toEqual([]);
  });

  it('结果文件夹只有经目录接口确认后才提供与执行打开动作', async () => {
    const f = sheetFixture();
    const v = files(f.app);
    f.book.SaveCopyAs = (path) => {
      v.data.set(path, 'PK\u0003\u0004copy');
    };
    const receipt = await f.host.apply(
      { type: 'document.saveCopy', fileName: '副本.xlsx', folder: 'C:\\out' },
      await f.host.readSelection(),
    );
    v.fs.readdirSync.mockReturnValue(undefined);
    await expect(f.host.openResult(receipt.actions![1]!.id)).rejects.toMatchObject({
      code: 'UNAVAILABLE',
    });
    expect(v.open).not.toHaveBeenCalled();
    f.app.FileSystem!.readdirSync = undefined;
    const next = await f.host.apply(
      { type: 'document.saveCopy', fileName: '另一副本.xlsx', folder: 'C:\\out' },
      await f.host.readSelection(),
    );
    expect(next.actions?.map((action) => action.label)).toEqual(['打开结果']);
  });

  it('通知正文兼容CRLF输入，不制造多余空行', async () => {
    const f = textFixture();
    await f.host.apply(
      { type: 'document.notice', fields: { ...fields, body: '第一行\r\n第二行' } },
      await f.host.readSelection(),
    );
    expect(f.target.Content!.Text).toContain('第一行\r第二行');
    expect(f.target.Content!.Text).not.toContain('第一行\r\r第二行');
  });

  it('原生Select被忽略时不报告定位成功', async () => {
    const f = sheetFixture();
    const before = await f.host.readSelection();
    const item = f.range.Cells!.Item;
    f.range.Cells = { Item: (row, column) => ({ ...item(row, column), Select: vi.fn() }) };
    await expect(f.host.locate(before, 1, 0)).rejects.toMatchObject({ code: 'UNAVAILABLE' });
  });
});

describe('WPS 页面设置与原生打印预览指引', () => {
  it('真实注册feature只请求打印设置，runner仍能预览、设置并恢复', async () => {
    const f = sheetFixture();
    const { runner } = createRuntime(f.host);
    const original = { ...f.page };
    const before = await f.preview();
    const read = vi.spyOn(f.host, 'readSelection');
    const task = await runner.prepare('sheet.print');
    expect(read).toHaveBeenCalledWith({ includePrintSettings: true });
    expect(task.changesDocument).toBe(true);
    expect(f.page).toEqual(original);
    expect(await runner.execute(task.id)).toMatchObject({
      outcome: 'completed',
      canUndo: true,
      message: expect.stringContaining('Ctrl+P'),
    });
    expect(f.page).toEqual({
      PrintArea: '$B$2:$C$4',
      Orientation: 2,
      Zoom: false,
      FitToPagesWide: 1,
      FitToPagesTall: false,
    });
    expect(runner.history).toHaveLength(1);
    await runner.undo();
    expect(f.page).toEqual(original);
    expect(await f.preview()).toEqual(before);
  });

  it.each([
    'cells',
    'areas',
    'wrap',
    'merge',
    'width',
    'bold',
    'position',
    'owner',
    'protection',
    'getter',
  ] as const)('打印缺少可靠布局%s时降级指引，显式布局读取仍拒绝', async (what) => {
    const f = sheetFixture();
    if (what === 'cells') f.range.Cells = undefined;
    if (what === 'areas') f.range.Areas = undefined;
    if (what === 'position') f.range.Row = undefined;
    if (what === 'owner') f.sheet.Parent = undefined;
    if (what === 'protection') f.sheet.ProtectContents = undefined;
    if (what === 'getter')
      Object.defineProperty(f.range, 'Cells', {
        get: () => {
          throw new Error('unsupported');
        },
      });
    if (what === 'wrap' || what === 'merge' || what === 'width' || what === 'bold') {
      const item = f.range.Cells!.Item;
      f.range.Cells!.Item = (row, column) => {
        const cell = item(row, column);
        Object.defineProperty(
          cell,
          { wrap: 'WrapText', merge: 'MergeCells', width: 'ColumnWidth', bold: 'Font' }[what]!,
          { value: undefined },
        );
        return cell;
      };
    }
    const original = { ...f.page };
    const snapshot = await f.host.readSelection({ includePrintSettings: true });
    expect(snapshot.printSettings).toMatchObject({ available: false, reason: expect.any(String) });
    expect(snapshot.sheetLayout).toBeUndefined();
    const explicit = expect(f.host.readSelection({ includeSheetLayout: true })).rejects;
    if (what === 'getter') await explicit.toThrow('unsupported');
    else await explicit.toMatchObject({ code: 'UNAVAILABLE' });
    await expect(
      f.host.readSelection({ includeSheetLayout: true, includePrintSettings: true }),
    ).rejects.toBeDefined();
    const { runner } = createRuntime(f.host);
    const task = await runner.prepare('sheet.print');
    expect(task.changesDocument).toBe(false);
    expect(await runner.execute(task.id)).toMatchObject({
      outcome: 'guidance',
      canUndo: false,
      message: expect.stringContaining('打印预览'),
    });
    expect(runner.history).toHaveLength(0);
    expect(f.page).toEqual(original);
    expect(f.state.writes).toEqual([]);
  });

  it('缺少固定范围接口时真实runner给出指引，不尝试写入', async () => {
    const f = sheetFixture();
    f.sheet.Range = undefined;
    const original = { ...f.page };
    const { runner } = createRuntime(f.host);
    const task = await runner.prepare('sheet.print');
    expect(task.changesDocument).toBe(false);
    expect(await runner.execute(task.id)).toMatchObject({ outcome: 'guidance', canUndo: false });
    expect(f.page).toEqual(original);
    expect(f.state.writes).toEqual([]);
  });

  it('打印预览后布局改变时真实runner拒绝过期设置', async () => {
    const f = sheetFixture();
    const { runner } = createRuntime(f.host);
    const task = await runner.prepare('sheet.print');
    f.state.widths[0] = 24;
    const original = { ...f.page };
    await expect(runner.execute(task.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(f.page).toEqual(original);
    expect(f.state.writes).toEqual([]);
  });

  it('当前选区、横向、宽一页高自动，完整读回并恢复原始设置', async () => {
    const f = sheetFixture();
    const original = { ...f.page };
    const receipt = await f.host.apply({ type: 'sheet.print' }, await f.printPreview());
    expect(f.page).toEqual({
      PrintArea: '$B$2:$C$4',
      Orientation: 2,
      Zoom: false,
      FitToPagesWide: 1,
      FitToPagesTall: false,
    });
    expect(receipt.message).toContain('Ctrl+P');
    f.state.widths[0] = 28;
    await f.host.undo(receipt);
    expect(f.page).toEqual(original);
    expect(f.state.widths[0]).toBe(28);
  });

  it('自动页高按公开0表示但原生false原样保存，未改属性保留用户后续修改', async () => {
    const f = sheetFixture();
    f.page.FitToPagesTall = false;
    expect((await f.printPreview()).printSettings!.fitTall).toBe(0);
    const receipt = await f.host.apply({ type: 'sheet.print' }, await f.printPreview());
    f.page.FitToPagesTall = 5;
    await f.host.undo(receipt);
    expect(f.page.FitToPagesTall).toBe(5);
    expect(f.page.Zoom).toBe(75);
  });

  it('被本次修改的打印属性发生冲突时不恢复', async () => {
    const f = sheetFixture();
    const receipt = await f.host.apply({ type: 'sheet.print' }, await f.printPreview());
    f.page.Orientation = 1;
    const current = { ...f.page };
    await expect(f.host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(f.page).toEqual(current);
  });

  it.each(['setup', 'enum'] as const)('缺少 %s 则available:false，返回未更改指引', async (what) => {
    const f = sheetFixture();
    if (what === 'setup') f.sheet.PageSetup = undefined;
    else f.app.Enum = undefined;
    const before = await f.printPreview();
    expect(before.printSettings).toMatchObject({ available: false, reason: expect.any(String) });
    expect(await f.host.apply({ type: 'sheet.print' }, before)).toMatchObject({
      outcome: 'guidance',
      canUndo: false,
    });
    expect(f.page.Zoom).toBe(75);
    expect(f.state.writes).toEqual([]);
  });
});

describe('WPS 保存副本、PDF与opaque结果动作', () => {
  it.each(['selection', 'activeSheet'] as const)(
    '%s 遇到工作表分组或未知单表状态时不弹保存窗口，不调用任何导出器',
    async (scope) => {
      for (const count of [undefined, 2]) {
        const f = pdfFixture();
        const partial = vi.fn();
        f.sheet.ExportAsFixedFormat = partial;
        (f.range as WpsSelection & WpsSheetPdfSource).ExportAsFixedFormat = partial;
        f.app.ActiveWindow!.SelectedSheets =
          count === undefined ? undefined : { Count: count, Item: () => f.sheet };
        const receipt = await f.host.apply(
          { type: 'document.exportPdf', fileName: '输出.pdf', scope },
          await f.host.readSelection(),
        );
        expect(receipt).toMatchObject({ outcome: 'guidance' });
        expect(partial).not.toHaveBeenCalled();
        expect(f.exportPdf).not.toHaveBeenCalled();
        expect(f.show).not.toHaveBeenCalled();
      }
    },
  );

  it('单表选择接口返回其它工作表或其它工作簿时拒绝导出', async () => {
    for (const selected of [
      { Name: '其它表', Parent: { Name: '费用.xlsx', FullName: 'C:\\work\\费用.xlsx' } },
      { Name: '费用', Parent: { Name: '其它.xlsx', FullName: 'C:\\work\\其它.xlsx' } },
    ]) {
      const f = pdfFixture();
      const partial = vi.fn();
      f.sheet.ExportAsFixedFormat = partial;
      f.app.ActiveWindow!.SelectedSheets = { Count: 1, Item: () => selected };
      expect(
        await f.host.apply(
          { type: 'document.exportPdf', fileName: '输出.pdf', scope: 'activeSheet' },
          await f.host.readSelection(),
        ),
      ).toMatchObject({ outcome: 'guidance' });
      expect(partial).not.toHaveBeenCalled();
      expect(f.show).not.toHaveBeenCalled();
    }
  });

  it('保存窗口期间工作表成组时停止导出，不取消分组', async () => {
    const f = pdfFixture();
    const partial = vi.fn();
    f.sheet.ExportAsFixedFormat = partial;
    f.show.mockImplementation(() => {
      f.app.ActiveWindow!.SelectedSheets!.Count = 2;
      return -1;
    });
    const receipt = await f.host.apply(
      { type: 'document.exportPdf', fileName: '输出.pdf', scope: 'activeSheet' },
      await f.host.readSelection(),
    );
    expect(receipt).toMatchObject({ outcome: 'guidance' });
    expect(partial).not.toHaveBeenCalled();
    expect(f.app.ActiveWindow!.SelectedSheets!.Count).toBe(2);
  });

  it('导出期间工作表成组时提示核对残件，不报告成功', async () => {
    const f = pdfFixture();
    f.sheet.ExportAsFixedFormat = vi.fn((_type, path) => {
      f.data.set(path, '%PDF-1.7\nselected scope\n%%EOF');
      f.app.ActiveWindow!.SelectedSheets!.Count = 2;
    });
    const { runner } = createRuntime(f.host);
    const receipt = await runner.execute(
      (
        await runner.prepare('document.exportPdf', {
          fileName: '输出.pdf',
          folder: 'C:\\out',
          exportScope: 'activeSheet',
        })
      ).id,
    );
    expect(receipt).toMatchObject({
      outcome: 'guidance',
      message: expect.stringContaining('不要直接重复导出'),
    });
    expect(receipt.actions).toBeUndefined();
    expect(runner.history).toHaveLength(0);
  });

  it.each(['unknown-story', 'other-story', 'unknown-enum', 'foreign-range'] as const)(
    '文字选区PDF的%s不调用保存窗口或任何导出器',
    async (mode) => {
      const f = textFixture();
      const v = files(f.app);
      const selected = vi.fn();
      (f.app.Selection as WpsWriterPdfSelection).ExportAsFixedFormat = selected;
      f.source.ExportAsFixedFormat = vi.fn();
      if (mode === 'unknown-story') f.app.Selection!.Range!.StoryType = undefined;
      if (mode === 'other-story') f.app.Selection!.Range!.StoryType = 42;
      if (mode === 'unknown-enum') f.app.Enum!.wdMainTextStory = undefined;
      if (mode === 'foreign-range') f.app.Selection!.Range!.Document = { Name: '其它.docx' };
      expect(
        await f.host.apply(
          { type: 'document.exportPdf', fileName: '输出.pdf', scope: 'selection' },
          await f.host.readSelection(),
        ),
      ).toMatchObject({ outcome: 'guidance', message: expect.stringContaining('正文区域') });
      expect(selected).not.toHaveBeenCalled();
      expect(f.source.ExportAsFixedFormat).not.toHaveBeenCalled();
      expect(v.fs.Exists).not.toHaveBeenCalled();
    },
  );

  it('文字导出期间切换文字区域时提示残件，不把同位置同文字当完成', async () => {
    const f = textFixture();
    const v = files(f.app);
    (f.app.Selection as WpsWriterPdfSelection).ExportAsFixedFormat = vi.fn((path) => {
      v.data.set(path, '%PDF-1.7\nselection\n%%EOF');
      f.app.Selection!.Range!.StoryType = 42;
    });
    const { runner } = createRuntime(f.host);
    const receipt = await runner.execute(
      (
        await runner.prepare('document.exportPdf', {
          fileName: '输出.pdf',
          folder: 'C:\\out',
          exportScope: 'selection',
        })
      ).id,
    );
    expect(receipt).toMatchObject({
      outcome: 'guidance',
      message: expect.stringContaining('不要直接重复导出'),
    });
    expect(receipt.actions).toBeUndefined();
    expect(runner.history).toHaveLength(0);
  });

  it('范围PDF的保存对话框期间数据变化时停止，不写文件也不退回整份导出', async () => {
    const f = pdfFixture();
    const partial = vi.fn();
    (f.range as WpsSelection & WpsSheetPdfSource).ExportAsFixedFormat = partial;
    f.show.mockImplementation(() => {
      f.state.values[1]![0] = '后来编辑';
      return -1;
    });
    const { runner } = createRuntime(f.host);
    const task = await runner.prepare('document.exportPdf', {
      fileName: '输出.pdf',
      exportScope: 'selection',
    });
    await expect(runner.execute(task.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(partial).not.toHaveBeenCalled();
    expect(f.exportPdf).not.toHaveBeenCalled();
    expect(f.data.has('C:\\out\\输出.pdf')).toBe(false);
  });

  it('范围PDF导出后数据变化保留残件提醒，不报告完成、不提供可打开的成功结果', async () => {
    const f = pdfFixture();
    (f.range as WpsSelection & WpsSheetPdfSource).ExportAsFixedFormat = vi.fn((_type, path) => {
      f.data.set(path, '%PDF-1.7\nselected scope\n%%EOF');
      f.state.values[1]![0] = '导出期间的编辑';
    });
    const { runner } = createRuntime(f.host);
    const receipt = await runner.execute(
      (
        await runner.prepare('document.exportPdf', {
          fileName: '输出.pdf',
          folder: 'C:\\out',
          exportScope: 'selection',
        })
      ).id,
    );
    expect(receipt).toMatchObject({
      outcome: 'guidance',
      message: expect.stringContaining('不要直接重复导出'),
    });
    expect(receipt.actions).toBeUndefined();
    expect(f.data.has('C:\\out\\输出.pdf')).toBe(true);
    expect(runner.history).toHaveLength(0);
  });

  it('文字选区接口缺失时不会调用整篇导出器', async () => {
    const f = textFixture();
    const v = files(f.app);
    f.source.ExportAsFixedFormat = vi.fn();
    expect(
      await f.host.apply(
        { type: 'document.exportPdf', fileName: '输出.pdf', folder: 'C:\\out', scope: 'selection' },
        await f.host.readSelection(),
      ),
    ).toMatchObject({ outcome: 'guidance', message: expect.stringContaining('仅当前选区') });
    expect(f.source.ExportAsFixedFormat).not.toHaveBeenCalled();
    expect(v.fs.Exists).not.toHaveBeenCalled();
  });

  it.each(['selection', 'activeSheet'] as const)(
    '所选范围 %s 仅调用原范围的导出器，忽略旧打印区、不触碰其它工作表',
    async (scope) => {
      const f = pdfFixture();
      // A hidden other sheet does not prohibit exporting this single scope.
      f.sheets[1]!.Visible = 48;
      const partial = vi.fn((_type: number, path: string) => {
        f.data.set(path, '%PDF-1.7\nselected scope\n%%EOF');
      });
      if (scope === 'selection')
        (f.range as WpsSelection & WpsSheetPdfSource).ExportAsFixedFormat = partial;
      else f.sheet.ExportAsFixedFormat = partial;
      const before = await f.host.readSelection();
      const settings = { ...f.page };
      const { runner } = createRuntime(f.host);
      const receipt = await runner.execute(
        (
          await runner.prepare('document.exportPdf', {
            fileName: '输出.pdf',
            folder: 'C:\\out',
            exportScope: scope,
          })
        ).id,
      );
      expect(receipt).toMatchObject({
        outcome: 'completed',
        message: expect.stringContaining(scope === 'selection' ? '仅当前选区' : '仅当前工作表'),
      });
      expect(partial).toHaveBeenCalledExactlyOnceWith(
        91,
        'C:\\out\\输出.pdf',
        undefined,
        undefined,
        true,
        undefined,
        undefined,
        false,
      );
      expect(f.exportPdf).not.toHaveBeenCalled();
      expect(f.item).not.toHaveBeenCalled();
      expect(await f.host.readSelection()).toEqual(before);
      expect(f.page).toEqual(settings);
    },
  );

  it.each(['selection', 'activeSheet'] as const)(
    '%s 接口缺失时不退回整个工作簿，也不弹保存窗口',
    async (scope) => {
      const f = pdfFixture();
      const receipt = await f.host.apply(
        { type: 'document.exportPdf', fileName: '输出.pdf', scope },
        await f.host.readSelection(),
      );
      expect(receipt).toMatchObject({ outcome: 'guidance' });
      expect(receipt.message).toContain(scope === 'selection' ? '仅当前选区' : '仅当前工作表');
      expect(f.exportPdf).not.toHaveBeenCalled();
      expect(f.show).not.toHaveBeenCalled();
    },
  );

  it('文字选区PDF传ExportCurrentPage=false，不导出包含未选文字的整页或整篇', async () => {
    const f = textFixture();
    const v = files(f.app);
    const selected = vi.fn((path: string) => {
      v.data.set(path, '%PDF-1.7\nselection\n%%EOF');
    });
    (f.app.Selection as WpsWriterPdfSelection).ExportAsFixedFormat = selected;
    f.source.ExportAsFixedFormat = vi.fn();
    const before = await f.host.readSelection();
    const receipt = await f.host.apply(
      { type: 'document.exportPdf', fileName: '输出.pdf', folder: 'C:\\out', scope: 'selection' },
      before,
    );
    expect(receipt).toMatchObject({
      outcome: 'completed',
      message: expect.stringContaining('仅当前选区'),
    });
    expect(selected).toHaveBeenCalledExactlyOnceWith(
      'C:\\out\\输出.pdf',
      17,
      false,
      undefined,
      false,
    );
    expect(f.source.ExportAsFixedFormat).not.toHaveBeenCalled();
    expect(await f.host.readSelection()).toEqual(before);
    expect(f.sourceWrites()).toBe(0);
  });

  it.each(['folder', 'picker', 'grouped-enums'] as const)(
    '可靠Sheets目录含可见图表页时真实runner导出整个工作簿（%s），忽略打印区域、不限制页码',
    async (mode) => {
      const f = pdfFixture();
      if (mode === 'grouped-enums') {
        f.app.Enum!.xlSheetVisible = undefined;
        f.app.Enum!.XlSheetVisibility = { xlSheetVisible: 47 };
        f.app.Enum!.xlTypePDF = undefined;
        f.app.Enum!.XlFixedFormatType = { xlTypePDF: 91 };
      }
      const { runner } = createRuntime(f.host);
      const before = await f.host.readSelection();
      const page = { ...f.page };
      const visibility = f.sheets.map((sheet) => sheet.Visible);
      const result = await runner.execute(
        (
          await runner.prepare('document.exportPdf', {
            fileName: '输出.pdf',
            ...(mode === 'picker' ? {} : { folder: 'C:\\out' }),
          })
        ).id,
      );
      expect(result).toMatchObject({
        outcome: 'completed',
        canUndo: false,
        location: 'C:\\out\\输出.pdf',
        message: expect.stringContaining('所有工作表'),
      });
      expect(f.exportPdf).toHaveBeenCalledExactlyOnceWith(
        91,
        'C:\\out\\输出.pdf',
        undefined,
        undefined,
        true,
        undefined,
        undefined,
        false,
      );
      expect(f.item).toHaveBeenCalledWith(3);
      expect(f.show).toHaveBeenCalledTimes(mode === 'picker' ? 1 : 0);
      expect(f.fs.readdirSync).toHaveBeenCalledWith('C:\\out');
      expect(f.fs.readAsBinaryString).toHaveBeenCalledWith('C:\\out\\输出.pdf');
      expect(f.sheets.map((sheet) => sheet.Visible)).toEqual(visibility);
      expect(f.page).toEqual(page);
      expect(await f.host.readSelection()).toEqual(before);
      expect(f.state.writes).toEqual([]);
      expect(runner.history).toHaveLength(1);
    },
  );

  it.each([
    'missing-sheets',
    'missing-item',
    'item-throws',
    'empty',
    'fractional',
    'unbounded',
    'duplicate',
    'foreign',
    'owner',
    'hidden',
    'very-hidden',
    'unknown',
    'chart-hidden',
    'visible-enum',
    'enum-nan',
    'directory-reader',
    'count-drift',
  ] as const)('PDF范围或目录不可靠（%s）时前置指引，不调用picker/导出/文件读取', async (what) => {
    const f = pdfFixture();
    if (what === 'missing-sheets') f.book.Sheets = undefined;
    if (what === 'missing-item') f.item.mockReturnValue(undefined!);
    if (what === 'item-throws')
      f.item.mockImplementation(() => {
        throw new Error('bridge failure');
      });
    if (what === 'empty' || what === 'fractional' || what === 'unbounded')
      Object.defineProperty(f.book.Sheets!, 'Count', {
        get: () => ({ empty: 0, fractional: 1.5, unbounded: 1001 })[what],
      });
    if (what === 'duplicate') f.sheets[1]!.Name = f.sheet.Name;
    if (what === 'foreign') f.sheets[1]!.Parent = { Name: '另一工作簿.xlsx' };
    if (what === 'owner') f.sheets[1]!.Parent = undefined;
    if (what === 'hidden' || what === 'very-hidden')
      f.sheets[1]!.Visible = what === 'hidden' ? 48 : 49;
    if (what === 'unknown') f.sheets[1]!.Visible = undefined;
    if (what === 'chart-hidden') f.sheets[2]!.Visible = 48;
    if (what === 'visible-enum') f.app.Enum!.xlSheetVisible = undefined;
    if (what === 'enum-nan') f.app.Enum!.xlSheetVisible = NaN;
    if (what === 'directory-reader') f.app.FileSystem!.readdirSync = undefined;
    if (what === 'count-drift') {
      let reads = 0;
      Object.defineProperty(f.book.Sheets!, 'Count', { get: () => (++reads === 1 ? 3 : 4) });
    }
    const visibility = f.sheets.map((sheet) => sheet.Visible);
    const { runner } = createRuntime(f.host);
    const result = await runner.execute(
      (await runner.prepare('document.exportPdf', { fileName: '输出.pdf' })).id,
    );
    expect(result).toMatchObject({
      outcome: 'guidance',
      canUndo: false,
      message: expect.stringContaining('所有工作表'),
    });
    expect(result.location).toBeUndefined();
    expect(result.actions).toBeUndefined();
    expect(f.app.FileDialog).not.toHaveBeenCalled();
    expect(f.exportPdf).not.toHaveBeenCalled();
    expect(f.fs.Exists).not.toHaveBeenCalled();
    expect(f.fs.readAsBinaryString).not.toHaveBeenCalled();
    expect(f.sheets.map((sheet) => sheet.Visible)).toEqual(visibility);
    expect(runner.history).toHaveLength(0);
  });

  it('仅有Worksheets无法排除未枚举的图表页，必须前置指引', async () => {
    const f = pdfFixture();
    f.book.Sheets = undefined;
    f.book.Worksheets = { Item: vi.fn(() => f.sheet) };
    const { runner } = createRuntime(f.host);
    expect(
      await runner.execute(
        (await runner.prepare('document.exportPdf', { fileName: '输出.pdf' })).id,
      ),
    ).toMatchObject({ outcome: 'guidance' });
    expect(f.book.Worksheets.Item).not.toHaveBeenCalled();
    expect(f.show).not.toHaveBeenCalled();
    expect(f.exportPdf).not.toHaveBeenCalled();
  });

  it.each(['missing', 'file', 'throws'] as const)('目标目录%s时不导出', async (what) => {
    const f = pdfFixture();
    if (what === 'missing') f.folders.delete('C:\\out');
    if (what === 'file') {
      f.folders.delete('C:\\out');
      f.data.set('C:\\out', 'not a directory');
    }
    if (what === 'throws')
      f.fs.readdirSync.mockImplementation(() => {
        throw new Error('not readable');
      });
    const { runner } = createRuntime(f.host);
    const result = await runner.execute(
      (await runner.prepare('document.exportPdf', { fileName: '输出.pdf', folder: 'C:\\out' })).id,
    );
    expect(result).toMatchObject({ outcome: 'guidance', canUndo: false });
    expect(f.show).not.toHaveBeenCalled();
    expect(f.exportPdf).not.toHaveBeenCalled();
    expect(f.fs.readAsBinaryString).not.toHaveBeenCalled();
  });

  it.each(['hidden', 'added', 'renamed'] as const)(
    '路径对话框返回后%s时重核对所有表，导出前拒绝',
    async (what) => {
      const f = pdfFixture();
      f.show.mockImplementation(() => {
        if (what === 'hidden') f.sheets[1]!.Visible = 48;
        if (what === 'added') f.sheets.push({ Name: '新表', Parent: f.book, Visible: 47 });
        if (what === 'renamed') f.sheets[1]!.Name = '新表名';
        return -1;
      });
      const { runner } = createRuntime(f.host);
      expect(
        await runner.execute(
          (await runner.prepare('document.exportPdf', { fileName: '输出.pdf' })).id,
        ),
      ).toMatchObject({ outcome: 'guidance', canUndo: false });
      expect(f.show).toHaveBeenCalledOnce();
      expect(f.exportPdf).not.toHaveBeenCalled();
    },
  );

  it('导出返回后可见状态变化不报告完成或发结果动作', async () => {
    const f = pdfFixture();
    f.exportPdf.mockImplementation((_type, path) => {
      f.data.set(path, '%PDF-1.7\nbytes\n%%EOF');
      f.sheets[1]!.Visible = 48;
    });
    const { runner } = createRuntime(f.host);
    const result = await runner.execute(
      (await runner.prepare('document.exportPdf', { fileName: '输出.pdf', folder: 'C:\\out' })).id,
    );
    expect(result).toMatchObject({ outcome: 'guidance', canUndo: false });
    expect(result.location).toBeUndefined();
    expect(result.actions).toBeUndefined();
    expect(result.message).toContain('可能已生成文件');
    expect(runner.history).toHaveLength(0);
  });

  it('ET SaveCopyAs到用户绝对路径，存在性及字节读回通过才发成功凭证', async () => {
    const f = sheetFixture();
    const v = files(f.app);
    f.book.SaveCopyAs = vi.fn((path) => {
      v.data.set(path, 'PK\u0003\u0004current-data');
    });
    const receipt = await f.host.apply(
      { type: 'document.saveCopy', fileName: '副本.xlsx', folder: 'C:\\out' },
      await f.host.readSelection(),
    );
    expect(receipt).toMatchObject({
      outcome: 'completed',
      canUndo: false,
      location: 'C:\\out\\副本.xlsx',
    });
    expect(f.book.FullName).toBe('C:\\work\\费用.xlsx');
    expect(f.book.SaveCopyAs).toHaveBeenCalledWith('C:\\out\\副本.xlsx');
    expect(receipt.actions).toHaveLength(2);
    await f.host.openResult(receipt.actions![0]!.id);
    await f.host.openResult(receipt.actions![1]!.id);
    expect(v.open.mock.calls).toEqual([
      ['C:\\out\\副本.xlsx', ''],
      ['C:\\out', ''],
    ]);
    await expect(f.host.openResult('C:\\out\\副本.xlsx')).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
    v.data.set('C:\\out\\副本.xlsx', 'changed');
    await expect(f.host.openResult(receipt.actions![0]!.id)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
  });

  it('原生对话框仅提供路径，显式SaveCopyAs执行；不调用Execute/SaveAs', async () => {
    const f = sheetFixture();
    const v = files(f.app);
    const execute = vi.fn();
    f.book.SaveCopyAs = vi.fn((path) => {
      v.data.set(path, 'PK\u0003\u0004copy');
    });
    const dialog = {
      Title: '',
      ButtonName: '',
      InitialFileName: '',
      Show: vi.fn(() => -1),
      SelectedItems: { Count: 1, Item: () => 'C:\\out\\副本.xlsx' },
      Execute: execute,
    };
    f.app.FileDialog = vi.fn(() => dialog);
    const receipt = await f.host.apply(
      { type: 'document.saveCopy', fileName: '副本.xlsx' },
      await f.host.readSelection(),
    );
    expect(receipt.outcome).toBe('completed');
    expect(dialog.Show).toHaveBeenCalledOnce();
    expect(execute).not.toHaveBeenCalled();
  });

  it('取消原生路径对话框不创建文件', async () => {
    const f = sheetFixture();
    files(f.app);
    f.book.SaveCopyAs = vi.fn();
    f.app.FileDialog = () => ({ InitialFileName: '', Title: '', ButtonName: '', Show: () => 0 });
    expect(
      await f.host.apply(
        { type: 'document.saveCopy', fileName: '副本.xlsx' },
        await f.host.readSelection(),
      ),
    ).toMatchObject({ outcome: 'cancelled', canUndo: false });
    expect(f.book.SaveCopyAs).not.toHaveBeenCalled();
  });

  it.each(['same', 'existing', 'relative', 'traversal', 'different-extension'] as const)(
    '%s 目标不得写入',
    async (what) => {
      const f = sheetFixture();
      const v = files(f.app);
      const save = vi.fn();
      f.book.SaveCopyAs = save;
      const plan: ExtendedActionPlan = {
        type: 'document.saveCopy',
        fileName:
          what === 'same' ? '费用.xlsx' : what === 'different-extension' ? '副本.xls' : '副本.xlsx',
        folder:
          what === 'same'
            ? 'C:\\work'
            : what === 'relative'
              ? 'out'
              : what === 'traversal'
                ? 'C:\\out\\..\\work'
                : 'C:\\out',
      };
      if (what === 'existing') v.data.set('C:\\out\\副本.xlsx', 'do-not-overwrite');
      const result = f.host.apply(plan, await f.host.readSelection());
      if (what === 'same' || what === 'existing')
        await expect(result).rejects.toMatchObject({ code: 'INVALID_PLAN' });
      else expect(await result).toMatchObject({ outcome: 'guidance' });
      expect(save).not.toHaveBeenCalled();
    },
  );

  it.each([undefined, 'C:\\out'])(
    'ET PDF无法保证所有工作表范围时前置指引（目录%s），不选路径、不导出',
    async (folder) => {
      const f = sheetFixture();
      const v = files(f.app);
      const exportPdf = vi.fn((_type: number, path: string) => {
        v.data.set(path, '%PDF-1.7\nonly visible print areas\n%%EOF\n');
      });
      f.book.ExportAsFixedFormat = exportPdf;
      const show = vi.fn(() => -1);
      f.app.FileDialog = vi.fn(() => ({
        InitialFileName: '',
        Title: '',
        ButtonName: '',
        Show: show,
        SelectedItems: { Count: 1, Item: () => 'C:\\out\\输出.pdf' },
      }));
      const { runner } = createRuntime(f.host);
      const task = await runner.prepare('document.exportPdf', { fileName: '输出.pdf', folder });
      expect(task.summary).toContain('所有工作表');
      expect(await runner.execute(task.id)).toMatchObject({
        outcome: 'guidance',
        canUndo: false,
        message: expect.stringContaining('所有工作表'),
      });
      expect(runner.history).toHaveLength(0);
      expect(f.app.FileDialog).not.toHaveBeenCalled();
      expect(show).not.toHaveBeenCalled();
      expect(exportPdf).not.toHaveBeenCalled();
      expect(v.fs.Exists).not.toHaveBeenCalled();
      expect(v.fs.readAsBinaryString).not.toHaveBeenCalled();
      expect(v.data.has('C:\\out\\输出.pdf')).toBe(false);
      expect(f.page.PrintArea).toBe('$A$1:$F$20');
    },
  );

  it.each(['missing-file', 'empty', 'not-pdf', 'throw', 'missing-read'] as const)(
    '%s 不得谎报PDF完成，也不发结果动作',
    async (what) => {
      const f = textFixture();
      const v = files(f.app);
      const exportPdf = vi.fn((path: string) => {
        if (what === 'throw') throw new Error('host failure');
        if (what !== 'missing-file') v.data.set(path, what === 'empty' ? '' : 'not a PDF');
      });
      f.source.ExportAsFixedFormat = exportPdf;
      if (what === 'missing-read') f.app.FileSystem!.readAsBinaryString = undefined;
      const receipt = await f.host.apply(
        { type: 'document.exportPdf', fileName: '输出.pdf', folder: 'C:\\out' },
        await f.host.readSelection(),
      );
      expect(receipt).toMatchObject({ outcome: 'guidance', canUndo: false });
      expect(receipt.location).toBeUndefined();
      expect(receipt.actions).toBeUndefined();
      if (what === 'missing-read') expect(exportPdf).not.toHaveBeenCalled();
    },
  );

  it('WPS文字PDF使用Document.ExportAsFixedFormat并保留原文档', async () => {
    const f = textFixture();
    const v = files(f.app);
    f.source.ExportAsFixedFormat = vi.fn((path) => {
      v.data.set(path, '%PDF-1.5\ndata\n%%EOF');
    });
    const receipt = await f.host.apply(
      { type: 'document.exportPdf', fileName: '通知.pdf', folder: 'C:\\out' },
      await f.host.readSelection(),
    );
    expect(receipt.outcome).toBe('completed');
    expect(f.source.ExportAsFixedFormat).toHaveBeenCalledWith('C:\\out\\通知.pdf', 17, false);
    expect(f.sourceWrites()).toBe(0);
  });

  it('文字副本只在Saved=true时复制磁盘字节；未保存改动直接指引', async () => {
    for (const saved of [true, false]) {
      const f = textFixture();
      const v = files(f.app);
      f.source.Saved = saved;
      v.data.set('C:\\work\\原稿.docx', 'PK\u0003\u0004saved-document');
      const receipt = await f.host.apply(
        { type: 'document.saveCopy', fileName: '副本.docx', folder: 'C:\\out' },
        await f.host.readSelection(),
      );
      expect(receipt.outcome).toBe(saved ? 'completed' : 'guidance');
      expect(v.fs.copyFileSync).toHaveBeenCalledTimes(saved ? 1 : 0);
      expect(f.sourceWrites()).toBe(0);
      expect(f.source.FullName).toBe('C:\\work\\原稿.docx');
    }
  });

  it.each(['saved', 'content', 'disk'] as const)(
    '文字复制后原稿%s变化时不报告成功',
    async (what) => {
      const f = textFixture();
      const v = files(f.app);
      v.data.set('C:\\work\\原稿.docx', 'PK\u0003\u0004original');
      v.fs.copyFileSync.mockImplementation((source, target) => {
        v.data.set(target, v.data.get(source)!);
        if (what === 'saved') f.source.Saved = false;
        if (what === 'content') f.source.Content!.Text = '用户在复制期间修改的内容';
        if (what === 'disk') v.data.set(source, 'changed-on-disk');
      });
      const receipt = await f.host.apply(
        { type: 'document.saveCopy', fileName: '副本.docx', folder: 'C:\\out' },
        await f.host.readSelection(),
      );
      expect(receipt.outcome).toBe('guidance');
      expect(receipt.actions).toBeUndefined();
    },
  );

  it.each(['string', 'typed-array'] as const)(
    '超128MiB的%s读回立即降级，不继续转换或比较',
    async (representation) => {
      const f = sheetFixture();
      const v = files(f.app);
      f.book.SaveCopyAs = (path) => {
        v.data.set(path, 'PK\u0003\u0004placeholder');
      };
      const length = 128 * 1024 * 1024 + 1;
      const raw =
        representation === 'string'
          ? 'x'.repeat(length)
          : new Proxy(new Uint8Array(1), {
              get(target, property) {
                if (property === 'length') return length;
                throw new Error(
                  `Must not convert oversized data: ${String(property)} ${target.byteLength}`,
                );
              },
            });
      f.app.FileSystem!.readAsBinaryString = () => raw;
      const receipt = await f.host.apply(
        { type: 'document.saveCopy', fileName: '副本.xlsx', folder: 'C:\\out' },
        await f.host.readSelection(),
      );
      expect(receipt.outcome).toBe('guidance');
      expect(receipt.actions).toBeUndefined();
    },
  );
});

describe('WPS新通知与原检查结果定位', () => {
  it.each(['documents', 'add', 'content'] as const)(
    '缺少新建通知接口%s时真实runner给出完整LF正文，并保留原稿恢复',
    async (what) => {
      const f = textFixture();
      const { runner } = createRuntime(f.host);
      const original = await f.host.readSelection();
      await runner.execute((await runner.prepare('text.enlarge')).id);
      if (what === 'documents') f.app.Documents = undefined;
      if (what === 'add') f.app.Documents = {} as NonNullable<WpsApplication['Documents']>;
      if (what === 'content') f.source.Content = undefined;
      const multiline = { ...fields, body: '第一段。\r\n第二段。\r第三段。\n末段。' };
      const result = await runner.execute(
        (await runner.prepare('document.notice', { notice: multiline })).id,
      );
      expect(result).toMatchObject({ outcome: 'guidance', canUndo: true });
      expect(result.message).toContain(noticeText(multiline).replace(/\r\n?/gu, '\n'));
      expect(result.message).not.toContain('\r');
      for (const value of [
        multiline.title,
        multiline.time,
        multiline.location,
        multiline.contact,
        '第一段。\n第二段。\n第三段。\n末段。',
      ])
        expect(result.message).toContain(value);
      expect(f.add).not.toHaveBeenCalled();
      expect(f.sourceWrites()).toBe(0);
      expect(runner.history).toHaveLength(1);
      await runner.undo();
      expect(await f.host.readSelection()).toEqual(original);
    },
  );

  it('真实runner原稿resize→notice→切回原稿→undo仍恢复原格式', async () => {
    const f = textFixture();
    const { runner } = createRuntime(f.host);
    const original = await f.host.readSelection();
    const selection = f.app.Selection;
    await runner.execute((await runner.prepare('text.enlarge')).id);
    expect((await f.host.readSelection()).fontSize).toBe(18);
    expect(
      await runner.execute((await runner.prepare('document.notice', { notice: fields })).id),
    ).toMatchObject({
      outcome: 'completed',
      canUndo: true,
    });
    expect(f.app.ActiveDocument).toBe(f.target);
    expect(runner.canUndo).toBe(true);
    const notification = f.target.Content!.Text;
    await expect(runner.undo()).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(runner.canUndo).toBe(true);
    f.app.ActiveDocument = f.source;
    f.app.Selection = selection;
    await runner.undo();
    expect(await f.host.readSelection()).toEqual(original);
    expect(runner.canUndo).toBe(false);
    expect(f.target.Content!.Text).toBe(notification);
    expect(f.sourceWrites()).toBe(0);
  });

  it('通知写入新建空白文档，原稿正文没有任何写入', async () => {
    const f = textFixture();
    const receipt = await f.host.apply(
      { type: 'document.notice', fields },
      await f.host.readSelection(),
    );
    expect(receipt).toMatchObject({ outcome: 'completed', canUndo: false });
    expect(receipt.message).toContain('未保存');
    expect(f.add).toHaveBeenCalledOnce();
    expect(f.target.Content!.Text.replace(/\r/gu, '\n').replace(/\n$/u, '')).toBe(
      noticeText(fields),
    );
    expect(f.sourceWrites()).toBe(0);
    expect(f.source.Content!.Text).toBe('原始正文\r');
  });

  it.each(['source', 'nonempty', 'wrong-owner'] as const)(
    '拒绝Add返回%s目标，不能写入原稿',
    async (what) => {
      const f = textFixture();
      if (what === 'source') f.app.Documents = { Add: () => f.source };
      if (what === 'nonempty') f.target.Content!.Text = '已有正文';
      if (what === 'wrong-owner') f.target.Content!.Document = f.source;
      await expect(
        f.host.apply({ type: 'document.notice', fields }, await f.host.readSelection()),
      ).rejects.toMatchObject({ code: 'INVALID_PLAN' });
      expect(f.sourceWrites()).toBe(0);
    },
  );

  it('Notice字段校验失败时不创建文档', async () => {
    const f = textFixture();
    await expect(
      f.host.apply(
        { type: 'document.notice', fields: { ...fields, title: '' } },
        await f.host.readSelection(),
      ),
    ).rejects.toMatchObject({ code: 'INVALID_PLAN' });
    expect(f.add).not.toHaveBeenCalled();
  });

  it('报告定位使用原区域内0-based坐标，移动选区后仍可定位且不写值', async () => {
    const f = sheetFixture();
    const original = await f.host.readSelection({ includeCellTypes: true });
    f.app.Selection = { ...f.range, Address: '$F$2:$G$4', Column: 6 };
    await f.host.locate(original, 2, 1);
    expect(f.state.selected).toEqual([3, 2]);
    expect(f.state.writes).toEqual([]);
    f.state.values[2]![1] = 999;
    await expect(f.host.locate(original, 1, 0)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(f.state.selected).toEqual([3, 2]);
  });

  it.each([
    [-1, 0],
    [0, -1],
    [3, 0],
    [0, 2],
    [0.5, 0],
  ])('拒绝越界/非整数定位%s,%s', async (r, c) => {
    const f = sheetFixture();
    await expect(f.host.locate(await f.host.readSelection(), r, c)).rejects.toMatchObject({
      code: 'INVALID_PLAN',
    });
    expect(f.state.selected).toEqual([]);
  });

  it('dispose同时撤销恢复凭证与结果动作', async () => {
    const f = sheetFixture();
    const receipt = await f.host.apply({ type: 'sheet.wrap' }, await f.preview());
    f.host.dispose();
    await expect(f.host.undo(receipt)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    await expect(f.host.openResult('opaque-id')).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
  });
});
