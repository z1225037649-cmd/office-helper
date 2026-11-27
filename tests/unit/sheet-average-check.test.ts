import { afterEach, describe, expect, it, vi } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import { OfficeAdapter } from '../../src/adapters/office/adapter';
import { WpsAdapter } from '../../src/adapters/wps/adapter';
import type { WpsApplication } from '../../src/adapters/wps/bridge';
import type { CellValue, HostAdapter } from '../../src/core/contracts';
import { checkSheetAverage } from '../../src/dev/sheet-average-check';

const originalData: CellValue[][] = [
  ['OFFICE_HELPER_NATIVE_TEST', '金额'],
  ['项目1', 120],
  ['项目2', 85],
  ['项目3', 45.5],
  ['项目4', 200],
];

function wpsFixture() {
  const state = {
    data: structuredClone(originalData),
    name: 'OfficeHelper-Native-Test.xlsx',
    address: '$B$2:$B$5',
  };
  const writes = vi.fn();
  const markerRead = vi.fn(() => structuredClone(state.data));
  const sheet = {
    Name: '费用',
    Range: vi.fn((address: string) => {
      if (address !== 'A1:B5') throw new Error('Unexpected native range');
      return {
        get Value2() {
          return markerRead();
        },
        set Value2(value: unknown) {
          writes(value);
          throw new Error('Must never write');
        },
      };
    }),
  };
  const isError = vi.fn((): boolean | number => false);
  const app: WpsApplication = {
    ActiveWorkbook: {
      get Name() {
        return state.name;
      },
      FullName: 'C:/fixtures/OfficeHelper-Native-Test.xlsx',
    },
    ActiveSheet: sheet,
    WorksheetFunction: { IsError: isError },
    Selection: {
      Rows: { Count: 4 },
      Columns: { Count: 1 },
      Areas: { Count: 1 },
      get Address() {
        return state.address;
      },
      get Value2() {
        return state.data.slice(1).map((row) => [row[1]!]);
      },
      set Value2(value: CellValue | CellValue[][] | undefined) {
        writes(value);
        throw new Error('Must never write selection');
      },
      Cells: {
        Item: (row: number, column: number) => {
          if (column !== 1) throw new Error('Unexpected column');
          return {
            get Value2() {
              return state.data[row]![1]!;
            },
          };
        },
      },
    },
  };
  vi.stubGlobal('window', { Application: app });
  const host = new WpsAdapter('sheet', () => app);
  const apply = vi.spyOn(host, 'apply');
  const undo = vi.spyOn(host, 'undo');
  const report = vi.fn(async (_stage: string, _details?: Record<string, unknown>) => undefined);
  return { host, state, writes, markerRead, report, apply, undo, isError, sheet };
}

function officeFixture() {
  const data = structuredClone(originalData);
  const writes = vi.fn();
  const selected = {
    address: '费用!B2:B5',
    rowCount: 4,
    columnCount: 1,
    get values() {
      return data.slice(1).map((row) => [row[1]!]);
    },
    set values(value: unknown) {
      writes(value);
      throw new Error('Must never write');
    },
    valueTypes: [['Integer'], ['Integer'], ['Double'], ['Integer']],
    load: vi.fn(),
  };
  const fixture = { values: data, load: vi.fn() };
  const sheet = {
    id: 'costs',
    load: vi.fn(),
    getRange: vi.fn((address: string) => {
      if (address !== 'A1:B5') throw new Error('Unexpected native range');
      return fixture;
    }),
  };
  const context = {
    workbook: {
      getSelectedRange: () => selected,
      worksheets: { getActiveWorksheet: () => sheet },
    },
    sync: vi.fn(async () => undefined),
  };
  const officeContext = {
    host: 'Excel',
    document: { url: 'C:/fixtures/OfficeHelper-Native-Test.xlsx' },
  };
  vi.stubGlobal('Office', { context: officeContext, HostType: { Excel: 'Excel' } });
  vi.stubGlobal('Excel', {
    run: vi.fn(async (read: (ctx: typeof context) => Promise<unknown>) => read(context)),
  });
  const host = new OfficeAdapter('sheet');
  const apply = vi.spyOn(host, 'apply');
  const undo = vi.spyOn(host, 'undo');
  const report = vi.fn(async (_stage: string, _details?: Record<string, unknown>) => undefined);
  return { host, writes, apply, undo, report, data, sheet, officeContext };
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.unstubAllEnvs();
});

describe('表格平均值开发验收守卫', () => {
  it('WPS 从预览取消、确认平均值到原有求和均只读，并报告专项通过', async () => {
    const { host, state, writes, report, apply, undo, markerRead } = wpsFixture();
    const before = structuredClone(state);
    await checkSheetAverage(host, report);
    expect(report.mock.calls.map(([stage]) => stage)).toEqual([
      'sheet-average-fixture-verified',
      'sheet-average-cancel-verified',
      'sheet-average-result-verified',
      'sheet-average-passed',
    ]);
    expect(report).toHaveBeenLastCalledWith('sheet-average-passed', {
      average: 112.625,
      total: 450.5,
      cells: 4,
      cancelled: true,
      canUndo: false,
      readonly: true,
    });
    expect(markerRead.mock.calls.length).toBeGreaterThan(10);
    expect(state).toEqual(before);
    expect(writes).not.toHaveBeenCalled();
    expect(apply).not.toHaveBeenCalled();
    expect(undo).not.toHaveBeenCalled();
  });

  it('Office 通过相同只读验收，使用原生A1:B5校验固定标记', async () => {
    const { host, writes, apply, undo, report, data, sheet } = officeFixture();
    await checkSheetAverage(host, report);
    expect(report.mock.calls.at(-1)?.[0]).toBe('sheet-average-passed');
    expect(data).toEqual(originalData);
    expect(sheet.getRange).toHaveBeenCalledWith('A1:B5');
    expect(writes).not.toHaveBeenCalled();
    expect(apply).not.toHaveBeenCalled();
    expect(undo).not.toHaveBeenCalled();
  });

  it('生产环境不允许运行专项', async () => {
    const { host, report, markerRead } = wpsFixture();
    vi.stubEnv('DEV', false);
    await expect(checkSheetAverage(host, report)).rejects.toThrow('开发环境');
    expect(report).not.toHaveBeenCalled();
    expect(markerRead).not.toHaveBeenCalled();
  });

  it('演示与伪造的宿主id不能冒充真实原生适配器', async () => {
    const host = new MockAdapter('sheet');
    host.info.id = 'wps';
    const report = vi.fn();
    await expect(checkSheetAverage(host, report)).rejects.toThrow(
      /真实 (?:Office\/WPS 表格适配器|原生表格适配器)/,
    );
    expect(report).not.toHaveBeenCalled();
  });

  it('自定义 HostAdapter 即使声明原生id也不能进入实机专项', async () => {
    const readSelection = vi.fn();
    const host = {
      info: {
        id: 'office' as const,
        label: '自定义适配器',
        kind: 'sheet' as const,
        capabilities: ['selection.read' as const],
      },
      readSelection,
      apply: vi.fn(),
      undo: vi.fn(),
      dispose: vi.fn(),
    } as unknown as HostAdapter;
    const report = vi.fn();
    await expect(checkSheetAverage(host, report)).rejects.toThrow(
      /真实 (?:Office\/WPS 表格适配器|原生表格适配器)/,
    );
    expect(readSelection).not.toHaveBeenCalled();
    expect(report).not.toHaveBeenCalled();
  });

  it('相同数值但错误文件名不允许验收', async () => {
    const { host, state, report, markerRead } = wpsFixture();
    state.name = '用户费用.xlsx';
    await expect(checkSheetAverage(host, report)).rejects.toThrow('专用测试文件');
    expect(report).not.toHaveBeenCalled();
    expect(markerRead).not.toHaveBeenCalled();
  });

  it.each(['$A$2:$A$5', '$B$7:$B$10', '$B$2:$B$5,$C$2'])(
    '错误选区 %s 不允许验收',
    async (address) => {
      const { host, state, report } = wpsFixture();
      state.address = address;
      await expect(checkSheetAverage(host, report)).rejects.toThrow('B2:B5');
      expect(report).not.toHaveBeenCalled();
    },
  );

  it('没有固定A1标记即使四个金额正确也拒绝运行', async () => {
    const { host, state, report } = wpsFixture();
    state.data[0]![0] = '用户工作表';
    await expect(checkSheetAverage(host, report)).rejects.toThrow('A1 标记');
    expect(report).not.toHaveBeenCalled();
  });

  it('缺失或未知类型信息不能绕过平均值检查', async () => {
    const { host, isError, report } = wpsFixture();
    isError.mockReturnValue(1);
    await expect(checkSheetAverage(host, report)).rejects.toThrow('原生类型');
    expect(report).not.toHaveBeenCalled();
  });

  it('验收期间非选区样本值改变时中止，绝不报告通过', async () => {
    const { host, state, report } = wpsFixture();
    report.mockImplementation(async (stage) => {
      if (stage === 'sheet-average-fixture-verified') state.data[0]![1] = '已修改';
    });
    await expect(checkSheetAverage(host, report)).rejects.toThrow('样本值');
    expect(report.mock.calls.map(([stage]) => stage)).not.toContain('sheet-average-passed');
  });

  it('Excel全局不是Excel宿主时拒绝通过Office适配器运行', async () => {
    const { host, officeContext, report } = officeFixture();
    officeContext.host = 'Word';
    await expect(checkSheetAverage(host, report)).rejects.toThrow('真实 Excel');
    expect(report).not.toHaveBeenCalled();
  });
});
