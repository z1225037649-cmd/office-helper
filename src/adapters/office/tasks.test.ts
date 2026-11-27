import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CellValue, ExtendedActionPlan, SelectionReadOptions } from '../../core/contracts';
import { FeatureRegistry } from '../../core/registry';
import { TaskRunner } from '../../core/runner';
import { features } from '../../features';
import { OfficeAdapter } from './adapter';

function excel(initialRows = 3) {
  const state = {
    values: [
      ['姓名', '电话', '备注'],
      ['甲', '123', ''],
      ['甲', '123', ''],
    ] as CellValue[][],
    types: [
      ['String', 'String', 'String'],
      ['String', 'String', 'Empty'],
      ['String', 'String', 'Empty'],
    ],
    widths: [42.5, 83.25, 109],
    wrap: [
      [false, true, false],
      [true, false, true],
      [false, false, true],
    ],
    bold: [
      [false, true, false],
      [true, false, true],
      [false, false, false],
    ],
    address: "'报名!表'!B4:D6",
    sheetId: 'sheet-original',
    sheetName: '报名!表',
    merged: false,
    protected: false,
    api13: true,
    api19: true,
    printArea: "'报名!表'!$A$1:$G$20",
    orientation: 'Portrait',
    zoom: { scale: 85, horizontalFitToPages: null, verticalFitToPages: null } as Record<
      string,
      number | null
    >,
    writes: [] as string[],
    selected: 'original',
    located: undefined as number[] | undefined,
    failAfterWrite: false,
    ignoreWrites: false,
  };
  if (initialRows !== 3) {
    state.values = Array.from({ length: initialRows }, () => ['值']);
    state.wrap = state.values.map(() => [false]);
    state.bold = state.values.map(() => [false]);
  }
  const document = { url: 'file:///报名.xlsx', getFileAsync: vi.fn() };
  const write = (property: string, change: () => void) => {
    state.writes.push(property);
    if (!state.ignoreWrites) change();
  };
  const pageLayout = {
    load: vi.fn(),
    get orientation() {
      return state.orientation;
    },
    set orientation(value: string) {
      write('orientation', () => {
        state.orientation = value;
      });
    },
    get zoom() {
      return structuredClone(state.zoom);
    },
    set zoom(value: Record<string, number | null>) {
      write('zoom', () => {
        state.zoom = structuredClone(value);
      });
    },
    getPrintAreaOrNullObject: () => ({
      get isNullObject() {
        return state.printArea === '';
      },
      get address() {
        return state.printArea;
      },
      load: vi.fn(),
    }),
    setPrintArea: (area: string | { address: string }) =>
      write('printArea', () => {
        state.printArea = typeof area === 'string' ? area : area.address;
      }),
  };
  const sheet = {
    get id() {
      return state.sheetId;
    },
    get name() {
      return state.sheetName;
    },
    load: vi.fn(),
    pageLayout,
    protection: {
      get protected() {
        return state.protected;
      },
      load: vi.fn(),
    },
    activate: vi.fn(),
    getRange: vi.fn(() => makeRange()),
  };
  function makeRange(
    row = 0,
    column = 0,
    rows = state.values.length,
    columns = state.values[0]!.length,
  ) {
    const each = (action: (r: number, c: number) => void) => {
      for (let r = row; r < row + rows; r++)
        for (let c = column; c < column + columns; c++) action(r, c);
    };
    const font = {
      load: vi.fn(),
      get bold() {
        return state.bold[row]?.[column];
      },
      set bold(value: boolean | undefined) {
        write('bold', () =>
          each((r, c) => {
            state.bold[r]![c] = value!;
          }),
        );
      },
    };
    const format = {
      load: vi.fn(),
      font,
      get columnWidth() {
        return state.widths[column];
      },
      set columnWidth(value: number | undefined) {
        write('columnWidth', () => {
          for (let c = column; c < column + columns; c++) state.widths[c] = value!;
        });
      },
      get wrapText() {
        return state.wrap[row]?.[column];
      },
      set wrapText(value: boolean | undefined) {
        write('wrapText', () =>
          each((r, c) => {
            state.wrap[r]![c] = value!;
          }),
        );
      },
      autofitColumns: vi.fn(() =>
        write('autofitColumns', () => {
          for (let c = column; c < column + columns; c++) state.widths[c] = 76 + c * 20;
        }),
      ),
    };
    const range = {
      get address() {
        return row === 0 &&
          column === 0 &&
          rows === state.values.length &&
          columns === state.values[0]!.length
          ? state.address
          : `other:${row}:${column}:${rows}:${columns}`;
      },
      get values() {
        return state.values.slice(row, row + rows).map((r) => r.slice(column, column + columns));
      },
      get valueTypes() {
        return state.types.slice(row, row + rows).map((r) => r.slice(column, column + columns));
      },
      get rowCount() {
        return rows;
      },
      get columnCount() {
        return columns;
      },
      get rowIndex() {
        return 3 + row;
      },
      get columnIndex() {
        return 1 + column;
      },
      worksheet: sheet,
      format,
      load: vi.fn(),
      track: vi.fn(() => range),
      untrack: vi.fn(),
      getColumn: (c: number) => makeRange(row, column + c, rows, 1),
      getRow: (r: number) => makeRange(row + r, column, 1, columns),
      getCell: (r: number, c: number) => makeRange(row + r, column + c, 1, 1),
      getMergedAreasOrNullObject: () => ({
        get isNullObject() {
          return !state.merged;
        },
        load: vi.fn(),
      }),
      select: vi.fn(() => {
        state.located = [3 + row, 1 + column];
        state.selected = 'located';
      }),
    };
    return range;
  }
  const original = makeRange();
  const selected = vi.fn(() => (state.selected === 'original' ? original : makeRange(1, 1, 1, 1)));
  const context = {
    workbook: {
      getSelectedRange: selected,
      worksheets: { getActiveWorksheet: () => sheet, getItem: vi.fn(() => sheet) },
    },
    sync: vi.fn(async () => {
      if (state.failAfterWrite && state.writes.length) {
        state.failAfterWrite = false;
        throw new Error('Native sync failed after mutation');
      }
    }),
  };
  vi.stubGlobal('Office', {
    context: {
      document,
      requirements: {
        isSetSupported: (set: string, version: string) =>
          set === 'ExcelApi' &&
          (version === '1.13' ? state.api13 : version === '1.9' ? state.api19 : true),
      },
    },
    FileType: { Pdf: 'pdf', Compressed: 'compressed' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
  });
  vi.stubGlobal('Excel', {
    run: vi.fn(async (...args: unknown[]) =>
      (args.at(-1) as (value: typeof context) => Promise<unknown>)(context),
    ),
  });
  return {
    host: new OfficeAdapter('sheet'),
    state,
    document,
    context,
    original,
    selected,
    sheet,
    makeRange,
  };
}

function word(api13 = true, hidden = false, canWrite = false) {
  const originalText = '原件中的内容，必须保留';
  const doc = { url: 'file:///原件.docx', getFileAsync: vi.fn() };
  const range = {
    text: originalText,
    font: { size: 12, bold: false, load: vi.fn() },
    paragraphs: {
      items: [
        {
          text: originalText,
          alignment: 'Left',
          lineSpacing: 12,
          firstLineIndent: 0,
          leftIndent: 0,
          rightIndent: 0,
          isListItem: false,
          load: vi.fn(),
        },
      ],
      load: vi.fn(),
    },
    load: vi.fn(),
    track: vi.fn(),
    untrack: vi.fn(),
    compareLocationWith: () => ({ value: 'Equal' }),
    getRange: vi.fn(),
    getOoxml: () => ({
      value:
        '<w:document><w:body><w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:p></w:body></w:document>',
    }),
  };
  const undoRange = {
    ...range,
    get text() {
      return range.text;
    },
    track: vi.fn(),
    untrack: vi.fn(),
  };
  range.track.mockReturnValue(range);
  undoRange.track.mockReturnValue(undoRange);
  range.getRange.mockReturnValue(undoRange);
  const createDocument = vi.fn((_base64: string) => created);
  const created = {
    open: vi.fn(),
    get body(): { text: string; load: ReturnType<typeof vi.fn> } {
      if (!hidden) throw new Error('WordApiHiddenDocument unavailable');
      return { text: '', load: vi.fn() };
    },
  };
  const context = {
    document: { getSelection: () => range },
    application: { createDocument },
    sync: vi.fn(async () => undefined),
  };
  vi.stubGlobal('Office', {
    context: {
      document: doc,
      requirements: {
        isSetSupported: (set: string) =>
          set === 'WordApi' ? api13 : set === 'WordApiHiddenDocument' ? hidden : false,
      },
    },
    FileType: { Pdf: 'pdf', Compressed: 'compressed' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
  });
  vi.stubGlobal('Word', {
    LocationRelation: { equal: 'Equal' },
    RangeLocation: { whole: 'Whole' },
    run: vi.fn(async (...args: unknown[]) =>
      (args.at(-1) as (ctx: typeof context) => unknown)(context),
    ),
  });
  return {
    host: new OfficeAdapter('text', canWrite),
    range,
    undoRange,
    doc,
    context,
    createDocument,
    created,
  };
}

function exportDouble(
  document: { getFileAsync: ReturnType<typeof vi.fn> },
  bytes = Array.from(new TextEncoder().encode('%PDF-1.7\nverified-pdf')),
) {
  const closeAsync = vi.fn((callback?: (value: unknown) => void) =>
    callback?.({ status: 'succeeded' }),
  );
  const file = {
    size: bytes.length,
    sliceCount: 2,
    closeAsync,
    getSliceAsync: vi.fn((index: number, callback: (value: unknown) => void) => {
      const data = index === 0 ? bytes.slice(0, 5) : bytes.slice(5);
      callback({ status: 'succeeded', value: { index, data, size: data.length } });
    }),
  };
  document.getFileAsync.mockImplementation(
    (_type: string, _options: unknown, callback: (result: unknown) => void) => {
      callback({ status: 'succeeded', value: file });
    },
  );
  let saved = new Uint8Array(0);
  const stream = {
    write: vi.fn(async (data: Uint8Array) => {
      saved = new Uint8Array(data);
    }),
    close: vi.fn(async () => undefined),
    abort: vi.fn(async () => undefined),
  };
  const handle = {
    name: '导出.pdf',
    createWritable: vi.fn(async () => stream),
    getFile: vi.fn(async () => ({ arrayBuffer: async () => saved.buffer })),
  };
  const picker = vi.fn(async () => handle);
  vi.stubGlobal('showSaveFilePicker', picker);
  return { file, stream, handle, picker, bytes };
}

function browserDownload() {
  const anchor = { href: '', download: '', hidden: false, click: vi.fn(), remove: vi.fn() };
  const document = { body: { appendChild: vi.fn() }, createElement: vi.fn(() => anchor) };
  vi.stubGlobal('showSaveFilePicker', undefined);
  vi.stubGlobal('document', document);
  const create = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:export');
  const revoke = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined);
  return { anchor, document, create, revoke };
}

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe('Office 表格原生任务：布局与按属性恢复', () => {
  it('仅按 opt-in 读取布局，保留原生磅数及每格混合值', async () => {
    const { host, state, original } = excel();
    expect(await host.readSelection()).not.toHaveProperty('sheetLayout');
    expect(original.getMergedAreasOrNullObject().isNullObject).toBe(true);
    const snapshot = await host.readSelection({ includeSheetLayout: true });
    expect(snapshot.sheetLayout).toEqual({
      columnWidths: [42.5, 83.25, 109],
      widthUnit: 'points',
      wrapText: state.wrap,
      headerBold: [false, true, false],
      merged: false,
      protected: false,
      firstRow: 3,
      firstColumn: 1,
      sheetName: '报名!表',
    });
    expect(state.writes).toEqual([]);
    expect(host.info.capabilities).toEqual(
      expect.arrayContaining([
        'sheet.autofit',
        'sheet.wrap',
        'sheet.header',
        'sheet.print',
        'selection.locate',
        'history.undo',
      ]),
    );
  });

  it.each(['sheet.autofit', 'sheet.wrap', 'sheet.header'] as const)(
    '%s 仅写目标属性，首行以原选区为准；换选区后可恢复混合原值',
    async (type) => {
      const { host, state } = excel();
      const before = structuredClone({ widths: state.widths, wrap: state.wrap, bold: state.bold });
      const expected = await host.readSelection({ includeSheetLayout: true });
      const receipt = await host.apply({ type }, structuredClone(expected));
      expect(receipt).toMatchObject({ canUndo: true, outcome: 'completed' });
      expect(
        state.writes.every(
          (write) =>
            write ===
            (type === 'sheet.autofit'
              ? 'autofitColumns'
              : type === 'sheet.wrap'
                ? 'wrapText'
                : 'bold'),
        ),
      ).toBe(true);
      if (type === 'sheet.header')
        expect(state.bold).toEqual([[true, true, true], ...before.bold.slice(1)]);
      if (type === 'sheet.wrap') expect(state.wrap.flat().every(Boolean)).toBe(true);
      state.selected = 'elsewhere';
      await host.undo(receipt);
      expect({ widths: state.widths, wrap: state.wrap, bold: state.bold }).toEqual(before);
      await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    },
  );

  it('wrap 恢复容许其他属性后续编辑，并保留后续列宽和首行粗体', async () => {
    const { host, state } = excel();
    const expected = await host.readSelection({
      includeSheetLayout: true,
      includePrintSettings: true,
    });
    const receipt = await host.apply({ type: 'sheet.wrap' }, expected);
    state.widths[0] = 222;
    state.bold[0]![0] = true;
    state.orientation = 'Landscape';
    await host.undo(receipt);
    expect(state.wrap).toEqual(expected.sheetLayout!.wrapText);
    expect(state.widths[0]).toBe(222);
    expect(state.bold[0]![0]).toBe(true);
    expect(state.orientation).toBe('Landscape');
  });

  it.each(['sheet.autofit', 'sheet.wrap', 'sheet.header'] as const)(
    '%s 目标属性后续被编辑时拒绝整个恢复',
    async (type) => {
      const { host, state } = excel();
      const receipt = await host.apply(
        { type },
        await host.readSelection({ includeSheetLayout: true }),
      );
      if (type === 'sheet.autofit') state.widths[1] = 999;
      if (type === 'sheet.wrap') state.wrap[1]![1] = false;
      if (type === 'sheet.header') state.bold[0]![1] = false;
      state.writes = [];
      await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(state.writes).toEqual([]);
    },
  );

  it.each(['merged', 'protected'] as const)('%s 预览或执行前均拒绝修改', async (flag) => {
    const { host, state } = excel();
    state[flag] = true;
    const expected = await host.readSelection({ includeSheetLayout: true });
    await expect(host.apply({ type: 'sheet.wrap' }, expected)).rejects.toMatchObject({
      code: 'UNAVAILABLE',
    });
    state[flag] = false;
    await expect(host.apply({ type: 'sheet.wrap' }, expected)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(state.writes).toEqual([]);
  });

  it.each(['data', 'layout', 'document', 'selection', 'range'] as const)(
    '%s 在预览后变化时不写入',
    async (change) => {
      const { host, state, document } = excel();
      const expected = await host.readSelection({
        includeSheetLayout: true,
        includePrintSettings: true,
      });
      if (change === 'data') state.values[1]![0] = '后续编辑';
      if (change === 'layout') state.wrap[1]![1] = true;
      if (change === 'document') document.url = 'file:///其他.xlsx';
      if (change === 'selection') state.selected = 'elsewhere';
      if (change === 'range') state.address = "'报名!表'!B5:D7";
      await expect(host.apply({ type: 'sheet.wrap' }, expected)).rejects.toMatchObject({
        code: 'STALE_SELECTION',
      });
      expect(state.writes).toEqual([]);
    },
  );

  it.each(['data', 'document', 'merged', 'protected'] as const)(
    '恢复前 %s 改变时不写入',
    async (change) => {
      const { host, state, document } = excel();
      const receipt = await host.apply(
        { type: 'sheet.header' },
        await host.readSelection({ includeSheetLayout: true }),
      );
      if (change === 'data') state.values[2]![0] = '改变';
      if (change === 'document') document.url = 'file:///其他.xlsx';
      if (change === 'merged' || change === 'protected') state[change] = true;
      state.writes = [];
      await expect(host.undo(receipt)).rejects.toMatchObject({
        code: change === 'merged' || change === 'protected' ? 'UNAVAILABLE' : 'STALE_SELECTION',
      });
      expect(state.writes).toEqual([]);
    },
  );

  it('原生部分失败及无效回读都不发成功凭据', async () => {
    const { host, state } = excel();
    let expected = await host.readSelection({ includeSheetLayout: true });
    state.failAfterWrite = true;
    await expect(host.apply({ type: 'sheet.wrap' }, expected)).rejects.toThrow('部分修改');
    expected = await host.readSelection({ includeSheetLayout: true });
    state.ignoreWrites = true;
    await expect(host.apply({ type: 'sheet.header' }, expected)).rejects.toThrow('结果检查');
  });

  it('5000×1 跨行分批读取而非每行同步', async () => {
    const { host, context } = excel(5000);
    const snapshot = await host.readSelection({ includeSheetLayout: true });
    expect(snapshot.sheetLayout!.wrapText).toHaveLength(5000);
    expect(context.sync.mock.calls.length).toBeLessThan(115);
  });

  it('缓存淘汰保留中的恢复 origin，后续 change 覆盖时释放这个孤立句柄', async () => {
    const { host, state, original, selected, makeRange } = excel();
    await host.apply(
      { type: 'sheet.wrap' },
      await host.readSelection({ includeSheetLayout: true }),
    );
    selected.mockImplementation(() => makeRange());
    for (let index = 0; index < 18; index++) {
      state.values[1]![0] = `新读数${index}`;
      await host.readSelection();
    }
    expect(original.untrack).not.toHaveBeenCalled();
    await host.apply(
      { type: 'sheet.header' },
      await host.readSelection({ includeSheetLayout: true }),
    );
    expect(original.untrack).toHaveBeenCalledOnce();
    host.dispose();
    expect(original.untrack).toHaveBeenCalledOnce();
  });
});

describe('Office 打印设置与原生确认边界', () => {
  it.each([false, true])('保留%s空打印区域和完整 zoom 原值，恢复不依赖当前选区', async (empty) => {
    const { host, state } = excel();
    if (empty) state.printArea = '';
    const before = structuredClone({
      area: state.printArea,
      orientation: state.orientation,
      zoom: state.zoom,
    });
    const expected = await host.readSelection({ includePrintSettings: true });
    const receipt = await host.apply({ type: 'sheet.print' }, expected);
    expect(state.printArea).toBe(state.address);
    expect(state.orientation).toBe('Landscape');
    expect(state.zoom).toEqual({ horizontalFitToPages: 1, verticalFitToPages: 0 });
    expect(receipt.message).toContain('Ctrl+P');
    expect(receipt.message).toContain('实际页数');
    expect(receipt.message).toContain('没有提交打印');
    expect(state.writes).toEqual(['printArea', 'orientation', 'zoom']);
    state.selected = 'elsewhere';
    state.wrap[0]![0] = true;
    await host.undo(receipt);
    expect({ area: state.printArea, orientation: state.orientation, zoom: state.zoom }).toEqual(
      before,
    );
    expect(state.wrap[0]![0]).toBe(true);
  });

  it.each(['printArea', 'orientation', 'zoom'] as const)(
    '打印 %s 后续编辑后拒绝恢复',
    async (property) => {
      const { host, state } = excel();
      const receipt = await host.apply(
        { type: 'sheet.print' },
        await host.readSelection({ includePrintSettings: true }),
      );
      if (property === 'printArea') state.printArea = 'A1:A5';
      if (property === 'orientation') state.orientation = 'Portrait';
      if (property === 'zoom') state.zoom = { scale: 90 };
      state.writes = [];
      await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(state.writes).toEqual([]);
    },
  );

  it('老版 pageLayout 不可用时给可执行的原生步骤', async () => {
    const { host, state } = excel();
    state.api19 = false;
    const expected = await host.readSelection({ includePrintSettings: true });
    expect(expected.printSettings?.available).toBe(false);
    expect(await host.apply({ type: 'sheet.print' }, expected)).toMatchObject({
      outcome: 'guidance',
      canUndo: false,
    });
    expect(state.writes).toEqual([]);
  });
});

describe('Office 检查报告定位绑定原始范围', () => {
  it('从原范围左上角计算偏移，不信任当前选区，且不改数据', async () => {
    const { host, state, selected, sheet } = excel();
    const snapshot = await host.readSelection({ includeCellTypes: true });
    state.selected = 'elsewhere';
    selected.mockClear();
    await host.locate(structuredClone(snapshot), 2, 1);
    expect(state.located).toEqual([5, 2]);
    expect(sheet.activate).toHaveBeenCalledOnce();
    expect(selected).not.toHaveBeenCalled();
    expect(state.writes).toEqual([]);
  });

  it.each(['data', 'types', 'document', 'sheet', 'address', 'forged'] as const)(
    '%s 变化或伪造后不定位',
    async (change) => {
      const { host, state, document } = excel();
      const snapshot = await host.readSelection({ includeCellTypes: true });
      if (change === 'data') state.values[1]![0] = '后续数据';
      if (change === 'types') state.types[1]![0] = 'Error';
      if (change === 'document') document.url = 'file:///其他.xlsx';
      if (change === 'sheet') state.sheetId = 'other-sheet';
      if (change === 'address') state.address = '其他!A1:C3';
      if (change === 'forged') snapshot.values![1]![0] = '伪造数据';
      await expect(host.locate(snapshot, 1, 0)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(state.located).toBeUndefined();
    },
  );

  it.each([
    [-1, 0],
    [3, 0],
    [0, 3],
    [0.5, 0],
    [0, NaN],
  ])('拒绝范围外偏移 %j', async (row, column) => {
    const { host, state } = excel();
    await expect(host.locate(await host.readSelection(), row!, column!)).rejects.toMatchObject({
      code: 'INVALID_PLAN',
    });
    expect(state.located).toBeUndefined();
  });
});

describe('Office 文件导出：saved/download/guidance 明确区分', () => {
  it.each(['selection', 'activeSheet'] as const)(
    'Excel %s PDF 明确指引对应范围，不调用整份导出',
    async (scope) => {
      const { host, document } = excel();
      const { picker } = exportDouble(document);
      const receipt = await host.apply(
        { type: 'document.exportPdf', fileName: '范围.pdf', scope },
        await host.readSelection(),
      );
      expect(receipt).toMatchObject({ outcome: 'guidance' });
      expect(receipt.message).toContain(scope === 'selection' ? '选定内容' : '活动工作表');
      expect(picker).not.toHaveBeenCalled();
      expect(document.getFileAsync).not.toHaveBeenCalled();
    },
  );

  it('Word 选区PDF不能静默导出整份文档', async () => {
    const { host, doc: document } = word();
    const { picker } = exportDouble(document);
    const receipt = await host.apply(
      { type: 'document.exportPdf', fileName: '范围.pdf', scope: 'selection' },
      await host.readSelection(),
    );
    expect(receipt).toMatchObject({ outcome: 'guidance' });
    expect(receipt.message).toContain('选项 → 选定内容');
    expect(picker).not.toHaveBeenCalled();
    expect(document.getFileAsync).not.toHaveBeenCalled();
  });

  it.each(['picker', 'download', 'folder'] as const)(
    'Excel PDF 无法保证整个工作簿范围时前置 guidance（%s），不导出或选择位置',
    async (mode) => {
      const { host, document, state } = excel();
      const { picker, file, stream } = exportDouble(document);
      const browser = mode === 'download' ? browserDownload() : undefined;
      const receipt = await host.apply(
        {
          type: 'document.exportPdf',
          fileName: '导出.pdf',
          ...(mode === 'folder' ? { folder: 'C:\\目标文件夹' } : {}),
        },
        await host.readSelection({ includeSheetLayout: true, includePrintSettings: true }),
      );
      expect(receipt).toMatchObject({ outcome: 'guidance', canUndo: false });
      expect(receipt.message).toContain('无法可靠确认');
      expect(receipt.message).toContain('助手未自动导出或保存');
      expect(receipt.message).toContain('选项 → 整个工作簿（所有工作表）');
      expect(receipt.message).toContain('导出.pdf');
      if (mode === 'folder') expect(receipt.message).toContain('C:\\目标文件夹');
      expect(receipt.location).toBeUndefined();
      expect(receipt.actions).toBeUndefined();
      expect(picker).not.toHaveBeenCalled();
      expect(document.getFileAsync).not.toHaveBeenCalled();
      expect(file.getSliceAsync).not.toHaveBeenCalled();
      expect(file.closeAsync).not.toHaveBeenCalled();
      expect(stream.write).not.toHaveBeenCalled();
      if (browser) {
        expect(browser.anchor.click).not.toHaveBeenCalled();
        expect(browser.create).not.toHaveBeenCalled();
      }
      expect(state.writes).toEqual([]);
      host.dispose();
    },
  );

  it('Word PDF 选择位置、原生分片、关闭 File、写入、回读全部成功才 completed', async () => {
    const { host, doc: document, range } = word();
    const original = range.text;
    const { file, stream, handle } = exportDouble(document);
    const receipt = await host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    expect(receipt).toMatchObject({ outcome: 'completed', canUndo: false, location: '导出.pdf' });
    expect(receipt.message).toContain('不向助手披露绝对路径');
    expect(file.getSliceAsync).toHaveBeenCalledTimes(2);
    expect(file.closeAsync).toHaveBeenCalledOnce();
    expect(stream.write).toHaveBeenCalledOnce();
    expect(stream.close).toHaveBeenCalledOnce();
    expect(handle.getFile).toHaveBeenCalledOnce();
    expect(range.text).toBe(original);
  });

  it('保存副本走 Compressed 原生文件，保持扩展名且不修改原件', async () => {
    const { host, document, state } = excel();
    const { handle } = exportDouble(document, [0x50, 0x4b, 3, 4, 0, 0, 1]);
    handle.name = '副本.xlsx';
    const receipt = await host.apply(
      { type: 'document.saveCopy', fileName: '副本.xlsx' },
      await host.readSelection(),
    );
    expect(receipt.outcome).toBe('completed');
    expect(document.getFileAsync.mock.calls[0]![0]).toBe('compressed');
    expect(state.writes).toEqual([]);
  });

  it('没有文件选择器时仅在成功原生导出后交浏览器，download 不宣称 saved', async () => {
    vi.useFakeTimers();
    const { host, doc: document } = word();
    const { file } = exportDouble(document);
    const { anchor, create, revoke } = browserDownload();
    const receipt = await host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    expect(receipt.outcome).toBe('download');
    expect(receipt.message).toContain('尚未确认文件已保存');
    expect(anchor.download).toBe('导出.pdf');
    expect(anchor.click).toHaveBeenCalledOnce();
    expect(create).toHaveBeenCalledOnce();
    expect(file.closeAsync).toHaveBeenCalledOnce();
    host.dispose();
    expect(revoke).toHaveBeenCalledWith('blob:export');
  });

  it('Excel 保存副本不受 PDF 范围限制，仍可原生导出后交浏览器下载', async () => {
    const { host, document, state } = excel();
    const { file } = exportDouble(document, [0x50, 0x4b, 3, 4, 0, 0, 1]);
    const { anchor } = browserDownload();
    const receipt = await host.apply(
      { type: 'document.saveCopy', fileName: '副本.xlsx' },
      await host.readSelection(),
    );
    expect(receipt.outcome).toBe('download');
    expect(document.getFileAsync.mock.calls[0]![0]).toBe('compressed');
    expect(file.closeAsync).toHaveBeenCalledOnce();
    expect(anchor.download).toBe('副本.xlsx');
    expect(anchor.click).toHaveBeenCalledOnce();
    expect(state.writes).toEqual([]);
    host.dispose();
  });

  it('不支持 export 时不自动启用下载，不创建成功结果', async () => {
    const { host, doc: document } = word();
    exportDouble(document);
    const { anchor, create } = browserDownload();
    document.getFileAsync.mockImplementation(
      (_t: string, _o: unknown, callback: (result: unknown) => void) =>
        callback({ status: 'failed', error: { message: 'NotSupported' } }),
    );
    const receipt = await host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    expect(receipt.outcome).toBe('guidance');
    expect(receipt.message).toContain('Word');
    expect(document.getFileAsync).toHaveBeenCalledOnce();
    expect(anchor.click).not.toHaveBeenCalled();
    expect(create).not.toHaveBeenCalled();
  });

  it.each(['readback', 'slice', 'signature', 'write', 'changed-document'] as const)(
    '%s 异常不报告保存成功，并关闭原生文件',
    async (failure) => {
      const { host, doc: document } = word();
      const { file, stream, handle } = exportDouble(document);
      if (failure === 'readback')
        handle.getFile.mockResolvedValue({ arrayBuffer: async () => new Uint8Array([0]).buffer });
      if (failure === 'slice')
        file.getSliceAsync.mockImplementation((_i, callback) => callback({ status: 'failed' }));
      if (failure === 'signature')
        file.getSliceAsync.mockImplementation((index, callback) =>
          callback({ status: 'succeeded', value: { index, data: [1], size: 1 } }),
        );
      if (failure === 'write') stream.write.mockRejectedValue(new Error('Disk full'));
      if (failure === 'changed-document')
        file.getSliceAsync.mockImplementation((index, callback) => {
          document.url = 'file:///other.docx';
          callback({
            status: 'succeeded',
            value: {
              index,
              data: index ? new Array(file.size - 5).fill(1) : [37, 80, 68, 70, 45],
              size: index ? file.size - 5 : 5,
            },
          });
        });
      const receipt = await host.apply(
        { type: 'document.exportPdf', fileName: '导出.pdf' },
        await host.readSelection(),
      );
      expect(receipt.outcome).toBe('guidance');
      expect(file.closeAsync).toHaveBeenCalledOnce();
      if (failure === 'write') expect(stream.abort).toHaveBeenCalledOnce();
      if (failure === 'changed-document') expect(stream.write).not.toHaveBeenCalled();
    },
  );

  it('用户取消选择位置返回 cancelled，未读取原生文件、未写文件', async () => {
    const { host, doc: document } = word();
    const { picker, stream } = exportDouble(document);
    picker.mockRejectedValue(Object.assign(new Error('Cancelled'), { name: 'AbortError' }));
    expect(
      (
        await host.apply(
          { type: 'document.exportPdf', fileName: '导出.pdf' },
          await host.readSelection(),
        )
      ).outcome,
    ).toBe('cancelled');
    expect(document.getFileAsync).not.toHaveBeenCalled();
    expect(stream.write).not.toHaveBeenCalled();
  });

  it.each(['folder', 'source', 'wrong-extension'] as const)(
    '%s 无法安全满足时回到原生步骤',
    async (condition) => {
      const { host, document } = excel();
      const { handle, stream } = exportDouble(
        document,
        condition === 'source' ? [0x50, 0x4b, 3, 4, 0, 0, 1] : undefined,
      );
      if (condition === 'source') handle.name = '报名.xlsx';
      const plan: ExtendedActionPlan =
        condition === 'folder'
          ? { type: 'document.exportPdf', fileName: '导出.pdf', folder: 'C:\\目标文件夹' }
          : {
              type: 'document.saveCopy',
              fileName: condition === 'source' ? '报名.xlsx' : '导出.docx',
            };
      const receipt = await host.apply(plan, await host.readSelection());
      expect(receipt.outcome).toBe('guidance');
      expect(stream.write).not.toHaveBeenCalled();
      if (condition === 'folder') expect(document.getFileAsync).not.toHaveBeenCalled();
    },
  );

  it('SharePoint URL 含查询参数时仍拒绝以原件文件名保存副本', async () => {
    const { host, document } = excel();
    document.url = 'https://example.test/报名.xlsx?web=1';
    const { picker, stream } = exportDouble(document);
    const receipt = await host.apply(
      { type: 'document.saveCopy', fileName: '报名.xlsx' },
      await host.readSelection(),
    );
    expect(receipt.outcome).toBe('guidance');
    expect(picker).not.toHaveBeenCalled();
    expect(stream.write).not.toHaveBeenCalled();
  });

  it('分片超时也关闭原生 File，失败后没有写入或结果操作', async () => {
    vi.useFakeTimers();
    const { host, doc: document } = word();
    const { file, stream } = exportDouble(document);
    file.getSliceAsync.mockImplementation(() => undefined);
    const pending = host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    await vi.advanceTimersByTimeAsync(15001);
    const receipt = await pending;
    expect(receipt.outcome).toBe('guidance');
    expect(receipt.actions).toBeUndefined();
    expect(file.closeAsync).toHaveBeenCalledOnce();
    expect(stream.write).not.toHaveBeenCalled();
  });

  it('拿文件超时之后的迟到回调仍释放 File，不触发写入', async () => {
    vi.useFakeTimers();
    const { host, doc: document } = word();
    const { file, stream } = exportDouble(document);
    let finish: ((response: unknown) => void) | undefined;
    document.getFileAsync.mockImplementation(
      (_t: string, _o: unknown, callback: (response: unknown) => void) => {
        finish = callback;
      },
    );
    const pending = host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    await vi.advanceTimersByTimeAsync(15001);
    expect((await pending).outcome).toBe('guidance');
    finish!({ status: 'succeeded', value: file });
    expect(file.closeAsync).toHaveBeenCalledOnce();
    expect(stream.write).not.toHaveBeenCalled();
  });
});

describe('Office 导出结果的不透明操作与资源释放', () => {
  it('只接受当前适配器返回的 action id，PDF 预览使用已验证的 Blob', async () => {
    const { host, doc: document } = word();
    exportDouble(document);
    const receipt = await host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    const action = receipt.actions![0]!;
    expect(action.label).toContain('PDF 预览');
    expect(action.id).not.toContain('导出.pdf');
    const { create, revoke } = browserDownload();
    const open = vi.fn();
    vi.stubGlobal('window', { open });
    await host.openResult(action.id);
    const blob = create.mock.calls[0]![0] as Blob;
    expect(await blob.text()).toBe('%PDF-1.7\nverified-pdf');
    expect(open).toHaveBeenCalledWith('blob:export', '_blank', 'noopener,noreferrer');
    await expect(host.openResult('C:\\伪造路径.pdf')).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
    await expect(host.openResult(receipt.token)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    host.dispose();
    expect(revoke).toHaveBeenCalledWith('blob:export');
    await expect(host.openResult(action.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
  });

  it('副本的操作明确为重新下载，不冒充打开原生文档或文件夹', async () => {
    const { host, document } = excel();
    const { handle } = exportDouble(document, [0x50, 0x4b, 3, 4, 0, 0, 1]);
    handle.name = '副本.xlsx';
    const receipt = await host.apply(
      { type: 'document.saveCopy', fileName: '副本.xlsx' },
      await host.readSelection(),
    );
    expect(receipt.actions![0]!.label).toContain('重新下载副本');
    const { anchor } = browserDownload();
    await host.openResult(receipt.actions![0]!.id);
    expect(anchor.download).toBe('副本.xlsx');
    expect(anchor.click).toHaveBeenCalledOnce();
    host.dispose();
  });

  it('浏览器下载返回的结果仍可预览，重复预览复用 URL', async () => {
    const { host, doc: document } = word();
    exportDouble(document);
    const { create } = browserDownload();
    const receipt = await host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    const open = vi.fn();
    vi.stubGlobal('window', { open });
    await host.openResult(receipt.actions![0]!.id);
    await host.openResult(receipt.actions![0]!.id);
    expect(create).toHaveBeenCalledOnce();
    expect(open).toHaveBeenCalledTimes(2);
    host.dispose();
  });

  it('结果归其他实例所有时拒绝打开', async () => {
    const { host, doc: document } = word();
    exportDouble(document);
    const receipt = await host.apply(
      { type: 'document.exportPdf', fileName: '导出.pdf' },
      await host.readSelection(),
    );
    const another = new OfficeAdapter('text');
    await expect(another.openResult(receipt.actions![0]!.id)).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
  });

  it('仅保留最近四个结果，淘汰时释放 URL 并使旧操作失效', async () => {
    const { host, doc: document } = word();
    exportDouble(document);
    const { revoke } = browserDownload();
    const snapshot = await host.readSelection();
    const receipts = [];
    for (let index = 0; index < 5; index++)
      receipts.push(
        await host.apply({ type: 'document.exportPdf', fileName: `导出${index}.pdf` }, snapshot),
      );
    await expect(host.openResult(receipts[0]!.actions![0]!.id)).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
    expect(revoke).toHaveBeenCalledOnce();
    host.dispose();
    expect(revoke).toHaveBeenCalledTimes(5);
  });
});

describe('Office 通知在新文档打开', () => {
  const fields = {
    title: '活动 <通知>',
    time: '2026年12月1日',
    location: '一楼 & 二楼',
    body: '请准时参加。\n携带材料。',
    contact: '工作人员',
  };
  it('WordApi 1.3 可用而 HiddenDocument 不可用时以完整 docx 创建，不改原件', async () => {
    const { host, range, createDocument, created } = word();
    const original = range.text;
    const receipt = await host.apply(
      { type: 'document.notice', fields },
      await host.readSelection(),
    );
    expect(receipt).toMatchObject({ outcome: 'completed', canUndo: false });
    expect(createDocument).toHaveBeenCalledOnce();
    expect(created.open).toHaveBeenCalledOnce();
    const base64 = createDocument.mock.calls[0]![0] as unknown as string;
    const zip = await JSZip.loadAsync(base64, { base64: true });
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('活动 &lt;通知&gt;');
    expect(xml).toContain('一楼 &amp; 二楼');
    expect(xml).toContain('携带材料。');
    expect(xml).toContain('联系人：工作人员');
    expect(range.text).toBe(original);
  });

  it('不支持 createDocument 时给出 Ctrl+N 和完整通知，不动原件', async () => {
    const { host, createDocument } = word(false);
    const receipt = await host.apply(
      { type: 'document.notice', fields },
      await host.readSelection(),
    );
    expect(receipt.outcome).toBe('guidance');
    expect(receipt.message).toContain('Ctrl+N');
    expect(receipt.message).toContain(fields.body);
    expect(createDocument).not.toHaveBeenCalled();
  });

  it('新文档 open 失败时返回可核对的指导，不声称创建完成', async () => {
    const { host, created } = word();
    created.open.mockImplementation(() => {
      throw new Error('Native open failed');
    });
    const receipt = await host.apply(
      { type: 'document.notice', fields },
      await host.readSelection(),
    );
    expect(receipt.outcome).toBe('guidance');
    expect(receipt.message).toContain('先检查是否已有新窗口');
  });

  it('通知预览后原件改变时拒绝创建', async () => {
    const { host, range, createDocument } = word();
    const expected = await host.readSelection();
    range.text = '后续编辑';
    await expect(host.apply({ type: 'document.notice', fields }, expected)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(createDocument).not.toHaveBeenCalled();
  });
});

// Exercise opt-in parity without relying on TaskRunner's separate preview guard.
describe('Office 执行入口 opt-in 字段与恢复凭据分流', () => {
  it.each<SelectionReadOptions>([
    { includeSheetLayout: true },
    { includeSheetLayout: true, includePrintSettings: true, includeCellTypes: true },
  ])('expected 携带 %j 时保持同组原生读取', async (options) => {
    const { host } = excel();
    const expected = await host.readSelection(options);
    const receipt = await host.apply({ type: 'sheet.header' }, expected);
    await expect(host.undo(receipt)).resolves.toBeUndefined();
  });
});

describe('Office 真实功能经 runner 执行', () => {
  const notice = {
    title: '活动通知',
    time: '2026年12月1日',
    location: '会议室',
    body: '请准时参加。\n携带材料。',
    contact: '工作人员',
  };

  it.each(['sheet.autofit', 'sheet.wrap', 'sheet.header', 'sheet.print'] as const)(
    '%s 经真实预览和执行后，换选区仍能恢复，并保留其他属性的后续编辑',
    async (id) => {
      const { host, state } = excel();
      const runner = new TaskRunner(host, new FeatureRegistry(features));
      const before = structuredClone({
        widths: state.widths,
        wrap: state.wrap,
        bold: state.bold,
        printArea: state.printArea,
        orientation: state.orientation,
        zoom: state.zoom,
      });
      const task = await runner.prepare(id);
      expect(state.writes).toEqual([]);
      expect(await runner.execute(task.id)).toMatchObject({
        outcome: 'completed',
        canUndo: true,
      });
      state.selected = 'elsewhere';
      state.bold[2]![2] = true;
      before.bold[2]![2] = true;
      if (id !== 'sheet.autofit') state.widths[0] = before.widths[0] = 333;
      await runner.undo();
      expect({
        widths: state.widths,
        wrap: state.wrap,
        bold: state.bold,
        printArea: state.printArea,
        orientation: state.orientation,
        zoom: state.zoom,
      }).toEqual(before);
      expect(runner.canUndo).toBe(false);
      expect(runner.history.map((entry) => entry.status)).toEqual(['restored', 'completed']);
      host.dispose();
    },
  );

  it('检查报告经真实类型读取后定位原范围，保留已有布局恢复凭据', async () => {
    const { host, state } = excel();
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    await runner.execute((await runner.prepare('sheet.wrap')).id);
    const task = await runner.prepare('sheet.checkRecords', {
      duplicates: {
        keyColumns: [0, 1],
        requiredColumns: [2],
        trimWhitespace: true,
        ignoreCase: false,
        hasHeader: true,
      },
    });
    const report = await runner.execute(task.id);
    const finding = report.findings![0]!;
    expect(finding).toBeDefined();
    expect(report.canUndo).toBe(true);
    state.selected = 'elsewhere';
    state.writes = [];
    await runner.locate(finding.id);
    expect(state.located).toEqual([3 + finding.row, 1 + finding.column]);
    expect(state.writes).toEqual([]);
    await runner.undo();
    expect(runner.canUndo).toBe(false);
    host.dispose();
  });

  it.each([false, true])('通知生成失败=%s 时原稿及旧格式 undo 均保留', async (fail) => {
    const { host, range, undoRange, created, createDocument } = word(true, false, true);
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    await runner.execute((await runner.prepare('text.enlarge')).id);
    const before = await host.readSelection();
    if (fail)
      created.open.mockImplementation(() => {
        throw new Error('Native open failed');
      });
    const task = await runner.prepare('document.notice', { notice });
    expect(createDocument).not.toHaveBeenCalled();
    expect(await runner.execute(task.id)).toMatchObject({
      outcome: fail ? 'guidance' : 'completed',
      canUndo: true,
    });
    expect(await host.readSelection()).toEqual(before);
    expect(undoRange.untrack).not.toHaveBeenCalled();
    await runner.undo();
    expect(range.font.size).toBe(12);
    expect(range.text).toBe(before.text);
    expect(undoRange.untrack).toHaveBeenCalledOnce();
    host.dispose();
  });

  it.each(['completed', 'cancelled', 'download', 'guidance'] as const)(
    'Word 导出实际返回 %s 时，历史状态准确且旧格式 undo 可用',
    async (outcome) => {
      const { host, doc, range } = word(true, false, true);
      const runner = new TaskRunner(host, new FeatureRegistry(features));
      await runner.execute((await runner.prepare('text.enlarge')).id);
      const { picker, stream } = exportDouble(doc);
      if (outcome === 'cancelled')
        picker.mockRejectedValue(Object.assign(new Error('Cancelled'), { name: 'AbortError' }));
      if (outcome === 'guidance')
        picker.mockRejectedValue(
          Object.assign(new Error('Permission denied'), { name: 'NotAllowedError' }),
        );
      const browser = outcome === 'download' ? browserDownload() : undefined;
      const task = await runner.prepare('document.exportPdf', { fileName: '导出.pdf' });
      const result = await runner.execute(task.id);
      expect(result).toMatchObject({ outcome, canUndo: true });
      expect(runner.history).toHaveLength(outcome === 'completed' ? 2 : 1);
      if (outcome === 'cancelled' || outcome === 'guidance') {
        expect(result.actions).toBeUndefined();
        expect(result.location).toBeUndefined();
        expect(doc.getFileAsync).not.toHaveBeenCalled();
        expect(stream.write).not.toHaveBeenCalled();
      }
      if (browser) expect(browser.anchor.click).toHaveBeenCalledOnce();
      await runner.undo();
      expect(range.font.size).toBe(12);
      host.dispose();
    },
  );

  it('Excel PDF 前置指导经 runner 保留旧布局 undo，不导出或弹保存窗口', async () => {
    const { host, document, state } = excel();
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    const before = structuredClone(state.wrap);
    await runner.execute((await runner.prepare('sheet.wrap')).id);
    const { picker } = exportDouble(document);
    const result = await runner.execute(
      (await runner.prepare('document.exportPdf', { fileName: '导出.pdf' })).id,
    );
    expect(result).toMatchObject({ outcome: 'guidance', canUndo: true });
    expect(result.actions).toBeUndefined();
    expect(picker).not.toHaveBeenCalled();
    expect(document.getFileAsync).not.toHaveBeenCalled();
    expect(runner.history).toHaveLength(1);
    await runner.undo();
    expect(state.wrap).toEqual(before);
    host.dispose();
  });

  it('真实导出结果只允许当前 runner 的授权 ID，新预览使旧按钮失效', async () => {
    const { host, doc } = word();
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    exportDouble(doc);
    const result = await runner.execute(
      (await runner.prepare('document.exportPdf', { fileName: '导出.pdf' })).id,
    );
    const action = result.actions![0]!;
    const { create, revoke } = browserDownload();
    const open = vi.fn();
    vi.stubGlobal('window', { open });
    await expect(runner.openResult('file:///other.pdf')).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
    expect(create).not.toHaveBeenCalled();
    await runner.openResult(action.id);
    expect(open).toHaveBeenCalledOnce();
    await runner.prepare('selection.inspect');
    await expect(runner.openResult(action.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(open).toHaveBeenCalledOnce();
    host.dispose();
    expect(revoke).toHaveBeenCalledWith('blob:export');
  });

  it.each([
    ['sheet', '.xlsx'],
    ['sheet', '.xlsm'],
    ['text', '.docx'],
  ] as const)(
    '%s %s 文件 URL 带查询和片段时，另存副本使用真实文件名和后缀',
    async (kind, extension) => {
      const setup = kind === 'sheet' ? excel() : word();
      const document = 'document' in setup ? setup.document : setup.doc;
      document.url = `https://example.test/${encodeURIComponent(`原件${extension}`)}?web=1#section`;
      const { handle, picker, stream } = exportDouble(document, [0x50, 0x4b, 3, 4, 0, 0, 1]);
      handle.name = `副本${extension}`;
      const runner = new TaskRunner(setup.host, new FeatureRegistry(features));

      expect((await setup.host.readSelection()).documentName).toBe(`原件${extension}`);
      const task = await runner.prepare('document.saveCopy', { fileName: handle.name });
      expect(stream.write).not.toHaveBeenCalled();
      expect(await runner.execute(task.id)).toMatchObject({
        outcome: 'completed',
        location: handle.name,
        canUndo: false,
      });
      expect(picker).toHaveBeenCalledOnce();
      expect(stream.write).toHaveBeenCalledOnce();
      expect(runner.history).toHaveLength(1);
      const staleTask = await runner.prepare('document.saveCopy', { fileName: handle.name });
      document.url = document.url.replace('web=1', 'web=2');
      await expect(runner.execute(staleTask.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(picker).toHaveBeenCalledOnce();
      expect(stream.write).toHaveBeenCalledOnce();
      setup.host.dispose();
    },
  );
});
