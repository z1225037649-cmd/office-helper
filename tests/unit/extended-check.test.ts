import { describe, expect, it, vi } from 'vitest';
import type {
  ActionPlan,
  ChangeReceipt,
  HostAdapter,
  SelectionReadOptions,
  SelectionSnapshot,
} from '../../src/core/contracts';
import { HelperError } from '../../src/core/errors';
import { noticeText } from '../../src/core/extended-actions';
import {
  assertNativeFixture,
  createExtendedObserver,
  type ExtendedNativeState,
} from '../../src/dev/extended-check-host';
import { ExtendedCheckSession } from '../../src/dev/extended-check-session';
import {
  assertExtendedConfig,
  assertExtendedSelection,
  assertSourcePath,
  expectedFindings,
  extendedRows,
  featureParameters,
  fixtureText,
  outputName,
  type ExtendedConfig,
} from '../../src/dev/extended-check-spec';

function config(kind: 'text' | 'sheet' = 'sheet', host: 'office' | 'wps' = 'wps'): ExtendedConfig {
  const run = 'a'.repeat(32);
  const matrix = `${host}-${kind}` as const;
  const base = `C:\\workspace\\output\\native\\extended\\${run}`;
  const fileName = `OfficeHelper-Extended-Test.${kind === 'text' ? 'docx' : 'xlsx'}`;
  return {
    schema: 1,
    run,
    matrix,
    host,
    kind,
    fileName,
    fixturePath: `${base}\\fixtures\\${matrix}\\${fileName}`,
    resultDirectory: `${base}\\results\\${matrix}`,
    text: fixtureText(run, matrix),
    values: structuredClone(extendedRows),
    notice: {
      title: `扩展诊断通知 ${run}`,
      time: '2026年11月28日 09:00（测试数据）',
      location: '隔离验收室（测试数据）',
      body: `OFFICE_HELPER_EXTENDED_TEST_V1\nRUN=${run}\nMATRIX=${matrix}`,
      contact: '测试联系人，不发送通知',
    },
  };
}

/** Unit double only; the real entry rejects this via createExtendedObserver. */
class UnitHost implements HostAdapter {
  readonly info;
  native: ExtendedNativeState;
  moved = false;
  print = {
    available: true,
    printArea: 'ExtendedData!$B$3:$D$6',
    orientation: 'portrait',
    zoom: 85 as number | false,
    fitWide: 2,
    fitTall: 3,
  };
  saved?: { native: ExtendedNativeState; print: UnitHost['print']; after?: string };
  outcome: ChangeReceipt['outcome'] = 'completed';
  notice = '';
  apply = vi.fn(async (plan: Exclude<ActionPlan, { type: 'report' }>): Promise<ChangeReceipt> => {
    if (plan.type.startsWith('document.')) {
      if (plan.type === 'document.notice') this.notice = noticeText(plan.fields);
      return {
        token: 'external',
        message: 'unit-only',
        canUndo: false,
        outcome: this.outcome,
        ...('fileName' in plan ? { location: `${plan.folder}\\${plan.fileName}` } : {}),
      };
    }
    this.saved = { native: structuredClone(this.native), print: structuredClone(this.print) };
    if (plan.type === 'sheet.autofit') this.native.widths = [11, 20, 21, 22, 23, 24];
    if (plan.type === 'sheet.wrap')
      for (let r = 2; r <= 9; r++) for (let c = 1; c <= 5; c++) this.native.wrap![r]![c] = true;
    if (plan.type === 'sheet.header') for (let c = 1; c <= 5; c++) this.native.bold![2]![c] = true;
    if (plan.type === 'sheet.print')
      this.print = {
        available: true,
        printArea: 'ExtendedData!$B$3:$F$10',
        orientation: 'landscape',
        zoom: false,
        fitWide: 1,
        fitTall: 0,
      };
    this.saved.after = JSON.stringify([this.native, this.print]);
    return { token: 'unit-change', message: 'unit-only', canUndo: true, outcome: this.outcome };
  });
  undo = vi.fn(async () => {
    if (!this.saved || JSON.stringify([this.native, this.print]) !== this.saved.after)
      throw new HelperError('STALE_SELECTION', 'unit conflict');
    this.native = structuredClone(this.saved.native);
    this.print = structuredClone(this.saved.print);
  });
  locate = vi.fn(async (_selection: SelectionSnapshot, row: number, column: number) => {
    this.moved = true;
    this.address = `${String.fromCharCode(66 + column)}${3 + row}`;
  });
  address = 'A2';
  constructor(readonly config: ExtendedConfig) {
    this.info = {
      id: config.host,
      kind: config.kind,
      label: 'unit-only',
      capabilities: [
        'selection.read',
        'sheet.autofit',
        'sheet.wrap',
        'sheet.header',
        'sheet.print',
        'selection.locate',
        'history.undo',
        'document.saveCopy',
        'document.exportPdf',
        'document.notice',
      ] as const,
    };
    if (config.kind === 'text') {
      this.native = { path: config.fixturePath, text: config.text };
      return;
    }
    const values: unknown[][] = Array.from({ length: 12 }, () => Array<unknown>(6).fill(''));
    values[0]!.splice(0, 3, 'OFFICE_HELPER_EXTENDED_TEST_V1', config.run, config.matrix);
    extendedRows.forEach((row, r) => values[r + 2]!.splice(1, 5, ...row));
    values[11]![1] = 'OFFICE_HELPER_EXTENDED_OUTSIDE';
    values[11]![5] = 909;
    const formulas = Array.from({ length: 12 }, () => Array<string | null>(6).fill(null));
    formulas[7]![5] = '=SUM(F4:F5)';
    this.native = {
      path: config.fixturePath,
      values,
      formulas,
      second: [['OFFICE_HELPER_EXTENDED_SECOND_SHEET', config.run, config.matrix]],
      catalogue: ['ExtendedData', 'ScopeSentinel'],
      widths: Array<number>(6).fill(11),
      bold: Array.from({ length: 12 }, () => Array<boolean>(6).fill(false)),
      wrap: Array.from({ length: 12 }, () => Array<boolean>(6).fill(false)),
      hidden: Array.from({ length: 12 }, (_, r) => r === 8),
    };
    this.native.bold![2]![1] = true;
  }
  async readSelection(options?: SelectionReadOptions): Promise<SelectionSnapshot> {
    const base = {
      documentId: 'unit-source',
      documentName: this.config.fileName,
      kind: this.config.kind,
      targetId: this.moved ? this.address : 'original',
      targetLabel: `ExtendedData!${this.moved ? this.address : '$B$3:$F$10'}`,
      revision: 'fixture',
    };
    if (this.config.kind === 'text')
      return { ...base, text: this.moved ? 'RUN' : this.config.text, fontSize: 12, bold: false };
    const values = this.moved ? [['']] : structuredClone(extendedRows);
    return {
      ...base,
      values,
      ...(options?.includeCellTypes
        ? {
            cellTypes: values.map((row) =>
              row.map((v) =>
                typeof v === 'string'
                  ? ('text' as const)
                  : typeof v === 'number'
                    ? ('number' as const)
                    : ('boolean' as const),
              ),
            ),
            rowScope: 'all' as const,
          }
        : {}),
      ...(options?.includeSheetLayout
        ? {
            sheetLayout: {
              columnWidths: this.native.widths!.slice(1),
              widthUnit: 'characters' as const,
              wrapText: this.native.wrap!.slice(2, 10).map((row) => row.slice(1)),
              headerBold: this.native.bold![2]!.slice(1),
              merged: false,
              protected: false,
              firstRow: 2,
              firstColumn: 1,
              sheetName: 'ExtendedData',
            },
          }
        : {}),
      ...(options?.includePrintSettings ? { printSettings: structuredClone(this.print) } : {}),
    };
  }
  dispose() {}
}
async function setup(kind: 'text' | 'sheet' = 'sheet', hostName: 'office' | 'wps' = 'wps') {
  const c = config(kind, hostName);
  const host = new UnitHost(c);
  const report = vi.fn(async (_stage: string, _details: Record<string, unknown>) => {});
  const local = vi.fn(async (_op: string) => ({ status: 'manual-required' }));
  const observer = {
    read: async () => structuredClone(host.native),
    notice: async () => ({
      source: structuredClone(host.native),
      text: host.notice,
      independent: true,
    }),
  };
  const session = new ExtendedCheckSession(host, c, observer, report, local);
  await session.initialize();
  return { c, host, report, local, session };
}

describe('扩展专项门禁和原生只读证据', () => {
  it.each(['office', 'wps'] as const)('严格接受 %s 对应矩阵和路径', (host) => {
    for (const kind of ['text', 'sheet'] as const) {
      const c = config(kind, host);
      expect(() => assertExtendedConfig(c)).not.toThrow();
      expect(() => assertSourcePath(c.fixturePath, c)).not.toThrow();
      expect(() =>
        assertSourcePath(`file:///${c.fixturePath.replaceAll('\\', '/')}`, c),
      ).not.toThrow();
    }
  });
  it.each([
    'matrix',
    'fileName',
    'text',
    'values',
    'resultDirectory',
    'fixturePath',
    'notice',
    'run',
  ])('拒绝被修改的配置 %s', (key) => {
    const c = { ...config(), [key]: key === 'values' ? [[1]] : 'wrong' };
    expect(() => assertExtendedConfig(c as ExtendedConfig)).toThrow();
  });
  it.each([
    'C:\\other\\OfficeHelper-Extended-Test.xlsx',
    'https://example.com/test.xlsx',
    'C:\\workspace\\..\\test.xlsx',
    '\\\\server\\share\\test.xlsx',
  ])('拒绝同名错路径/远端/遍历 %s', (path) =>
    expect(() => assertSourcePath(path, config())).toThrow(),
  );
  it('正文不能仅含前缀；矩阵不能仅含部分匹配值', async () => {
    const { c, host } = await setup();
    const selection = await host.readSelection();
    expect(() =>
      assertExtendedSelection({ ...selection, documentName: 'other.xlsx' }, c),
    ).toThrow();
    expect(() => assertExtendedSelection({ ...selection, values: [[1]] }, c)).toThrow();
    const text = config('text');
    expect(() =>
      assertExtendedSelection(
        { ...selection, kind: 'text', documentName: text.fileName, text: `${text.text}\nextra` },
        text,
      ),
    ).toThrow();
  });
  it('真实入口拒绝单测替身，不因 host.info.id 而信任', () => {
    const c = config();
    expect(() => createExtendedObserver(new UnitHost(c), c)).toThrow('不接受 mock');
  });
  it.each(['formulas', 'second', 'catalogue', 'hidden'] as const)(
    '拒绝缺少原生 %s 证据',
    async (key) => {
      const { host, c } = await setup();
      delete host.native[key];
      expect(() => assertNativeFixture(host.native, c)).toThrow();
    },
  );
  it('导出参数固定本轮目录和唯一名称，不能向文件选择器开放任意目录', () => {
    const c = config();
    expect(featureParameters('document.exportPdf', c)).toEqual({
      folder: c.resultDirectory,
      fileName: outputName('document.exportPdf', c),
    });
  });
});

describe('真实 runtime runner 编排单测（不是实机验收）', () => {
  it.each([
    'sheet.autofit',
    'sheet.wrap',
    'sheet.header',
    'sheet.print',
    'sheet.checkRecords',
    'document.saveCopy',
    'document.exportPdf',
  ] as const)('%s 取消与替换预览均不调用写入', async (feature) => {
    const { host, session } = await setup();
    await session.prepare(feature);
    await session.cancel();
    await session.prepare(feature);
    await session.supersede();
    expect(host.apply).not.toHaveBeenCalled();
    expect(session.evidence).toContainEqual({ check: `${feature}.expired`, status: 'passed' });
  });
  it.each(['sheet.autofit', 'sheet.wrap', 'sheet.header', 'sheet.print'] as const)(
    '%s 确认读回、移动选区拒绝旧预览、恢复及冲突',
    async (feature) => {
      const { host, session } = await setup();
      const before = structuredClone(host.native);
      await session.prepare(feature);
      host.moved = true;
      await session.stale();
      expect(host.apply).not.toHaveBeenCalled();
      host.moved = false;
      await session.prepare(feature);
      await session.confirm();
      await session.undo();
      expect(host.native).toEqual(before);
      await session.prepare(feature);
      await session.confirm();
      const knownAfter = structuredClone([host.native, host.print]);
      if (feature === 'sheet.print') host.print.orientation = 'portrait';
      else if (feature === 'sheet.autofit') host.native.widths![1] = 999;
      else if (feature === 'sheet.wrap') host.native.wrap![3]![2] = false;
      else host.native.bold![2]![2] = false;
      const conflictState = structuredClone([host.native, host.print]);
      await session.conflict();
      expect([host.native, host.print]).toEqual(conflictState);
      [host.native, host.print] = knownAfter as [ExtendedNativeState, typeof host.print];
      await session.undo();
      expect(host.native).toEqual(before);
    },
  );
  it('复用本地 claim，抢占失败前没有 runner 写入', async () => {
    const c = config();
    const host = new UnitHost(c);
    const session = new ExtendedCheckSession(
      host,
      c,
      { read: async () => host.native, notice: vi.fn() },
      async () => {
        throw new Error('409 claim');
      },
      vi.fn(),
    );
    await expect(session.initialize()).rejects.toThrow('409');
    await expect(session.action(() => session.prepare('sheet.wrap'))).rejects.toThrow('尚未核验');
    expect(host.apply).not.toHaveBeenCalled();
  });
  it('缺少接口只记 unsupported，guidance/download/cancelled 不记 passed', async () => {
    for (const outcome of ['guidance', 'download', 'cancelled'] as const) {
      const { host, session } = await setup();
      host.outcome = outcome;
      await session.prepare('document.exportPdf');
      await session.confirm();
      expect(session.evidence).toContainEqual({
        check: 'document.exportPdf.confirm',
        status: 'unsupported',
      });
      expect(session.evidence).not.toContainEqual({
        check: 'document.exportPdf.confirm',
        status: 'passed',
      });
    }
  });
  it('导出确认前 reserve，落盘只读核对后仍需要全范围/位置人工验证', async () => {
    const { host, session, local } = await setup();
    await session.prepare('document.saveCopy');
    await session.confirm();
    expect(local.mock.calls.map(([op]) => op)).toEqual(['reserve-copy', 'verify-copy']);
    expect(host.apply).toHaveBeenCalledOnce();
    expect(session.evidence).toContainEqual({
      check: 'document.saveCopy.confirm',
      status: 'manual-required',
    });
  });
  it('reserve 拒绝后不能执行任何宿主导出', async () => {
    const { host, session, local } = await setup();
    local.mockRejectedValueOnce(new Error('exists'));
    await session.prepare('document.exportPdf');
    await expect(session.action(() => session.confirm())).rejects.toThrow('exists');
    expect(host.apply).not.toHaveBeenCalled();
    expect(session.state.stopped).toBe(true);
  });
  it('查重包含隐藏行，逐个定位到原始偏移坐标', async () => {
    const { host, session } = await setup();
    await session.prepare('sheet.checkRecords');
    await session.confirm();
    expect(session.state.findings).toHaveLength(4);
    for (const finding of expectedFindings) {
      await session.locate(finding.id);
      expect(host.address).toBe(finding.address);
    }
    expect(host.apply).not.toHaveBeenCalled();
  });
  it('WPS 通知确认须读回独立未保存文档和全部正文；取消没有新建调用', async () => {
    const { host, session } = await setup('text');
    await session.prepare('document.notice');
    await session.cancel();
    expect(host.apply).not.toHaveBeenCalled();
    await session.prepare('document.notice');
    await session.confirm();
    expect(session.evidence).toContainEqual({ check: 'document.notice.confirm', status: 'passed' });
  });
  it('Office 通知即使接口 completed 也不能伪造新窗口读回通过', async () => {
    const { session } = await setup('text', 'office');
    await session.prepare('document.notice');
    await session.confirm();
    expect(session.evidence).toContainEqual({
      check: 'document.notice.confirm',
      status: 'manual-required',
    });
  });
  it('人工证据和总报告不会把 unsupported 提升为 passed', async () => {
    const { host, session, report } = await setup();
    host.outcome = 'guidance';
    await session.prepare('document.exportPdf');
    await session.confirm();
    await session.manual(
      'output-scope-location',
      'WPS 测试版本；已人工打开全工作簿，截图位于隔离证据目录。',
    );
    await session.summary();
    expect(session.evidence).toContainEqual({
      check: 'document.exportPdf.confirm',
      status: 'unsupported',
    });
    expect(report.mock.calls.at(-1)?.[1]).toMatchObject({ acceptance: 'incomplete' });
  });
  it('原稿或公式错误会在 prepare 前失败，不能调用宿主修改', async () => {
    const { host, session } = await setup();
    host.native.formulas![7]![5] = null;
    await expect(session.action(() => session.prepare('sheet.wrap'))).rejects.toThrow('公式');
    expect(host.apply).not.toHaveBeenCalled();
    expect(session.state.stopped).toBe(true);
  });
});
