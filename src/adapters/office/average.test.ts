import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CellValue, FeatureDefinition } from '../../core/contracts';
import { FeatureRegistry } from '../../core/registry';
import { TaskRunner } from '../../core/runner';
import { OfficeAdapter } from './adapter';

function setupExcel() {
  const state: {
    values: CellValue[][];
    valueTypes: unknown;
    address: string;
    rowCount: number;
    columnCount: number;
  } = {
    values: [[120], [85], [45.5], [200]],
    valueTypes: [['Integer'], ['Integer'], ['Double'], ['Integer']],
    address: '费用!B2:B5',
    rowCount: 4,
    columnCount: 1,
  };
  const writes = vi.fn();
  const valuesRead = vi.fn();
  const typesRead = vi.fn();
  const load = vi.fn();
  const range = {
    get address() {
      return state.address;
    },
    get rowCount() {
      return state.rowCount;
    },
    get columnCount() {
      return state.columnCount;
    },
    get values() {
      valuesRead();
      return state.values;
    },
    set values(value: CellValue[][]) {
      writes(value);
    },
    get valueTypes() {
      typesRead();
      return state.valueTypes;
    },
    load,
  };
  const sheet = { id: 'costs', load: vi.fn() };
  const context = {
    workbook: {
      getSelectedRange: () => range,
      worksheets: { getActiveWorksheet: () => sheet },
    },
    sync: vi.fn(async () => undefined),
  };
  vi.stubGlobal('Office', { context: { document: { url: 'C:/费用.xlsx' } } });
  vi.stubGlobal('Excel', {
    run: vi.fn(async (read: (ctx: typeof context) => Promise<unknown>) => read(context)),
  });
  const host = new OfficeAdapter('sheet');
  return { host, state, range, context, load, writes, valuesRead, typesRead };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('Excel 平均值按需类型读取', () => {
  it('值和原生类型在同一次同步读取，完整范围包含隐藏及筛选行', async () => {
    const { host, load, context, writes } = setupExcel();
    const snapshot = await host.readSelection({ includeCellTypes: true });

    expect(load.mock.calls).toEqual([
      [['address', 'rowCount', 'columnCount']],
      [['values', 'valueTypes']],
    ]);
    expect(context.sync).toHaveBeenCalledTimes(2);
    expect(snapshot).toMatchObject({
      values: [[120], [85], [45.5], [200]],
      cellTypes: [['number'], ['number'], ['number'], ['number']],
      rowScope: 'all',
      targetId: 'costs:费用!B2:B5',
    });
    expect(writes).not.toHaveBeenCalled();
    expect(host.info.capabilities).toEqual(['selection.read']);
  });

  it('明确区分错误、错误样式文本、空白、布尔和不认识的类型', async () => {
    const { host, state } = setupExcel();
    state.values = [[null, '#VALUE!', '#VALUE!', true, 0, 1.5, { rich: true } as never]];
    state.valueTypes = [['Empty', 'Error', 'String', 'Boolean', 'Integer', 'Double', 'RichValue']];
    state.rowCount = 1;
    state.columnCount = 7;

    expect((await host.readSelection({ includeCellTypes: true })).cellTypes).toEqual([
      ['empty', 'error', 'text', 'boolean', 'number', 'number', 'unknown'],
    ]);
  });

  it.each([undefined, null, 'Integer', ['Integer'], [null]])(
    '缺失或非矩阵类型 %j 不冒充可靠类型信息',
    async (valueTypes) => {
      const { host, state } = setupExcel();
      state.valueTypes = valueTypes;

      expect((await host.readSelection({ includeCellTypes: true })).cellTypes).toBeUndefined();
    },
  );

  it('原有合计和查看范围无需读取类型，兼容未提供类型的宿主', async () => {
    const { host, state, load, typesRead, writes } = setupExcel();
    state.valueTypes = undefined;

    const snapshot = await host.readSelection();

    expect(load.mock.calls).toEqual([[['address', 'rowCount', 'columnCount']], ['values']]);
    expect(snapshot.values).toEqual(state.values);
    expect(snapshot).not.toHaveProperty('cellTypes');
    expect(snapshot).not.toHaveProperty('rowScope');
    expect(typesRead).not.toHaveBeenCalled();
    expect(writes).not.toHaveBeenCalled();
  });

  it('超限时先拒绝，再不加载或读取任何值及类型', async () => {
    const { host, state, load, valuesRead, typesRead } = setupExcel();
    state.rowCount = 5001;

    await expect(host.readSelection({ includeCellTypes: true })).rejects.toMatchObject({
      code: 'SELECTION_TOO_LARGE',
    });
    expect(load.mock.calls).toEqual([[['address', 'rowCount', 'columnCount']]]);
    expect(valuesRead).not.toHaveBeenCalled();
    expect(typesRead).not.toHaveBeenCalled();
  });

  it('值未变而类型由文本变成错误时，确认前拒绝旧报告', async () => {
    const { host, state, writes } = setupExcel();
    state.values = [['#VALUE!']];
    state.valueTypes = [['String']];
    state.rowCount = 1;
    const detailedReport: FeatureDefinition = {
      id: 'test.detailed-report',
      title: '类型报告',
      description: '读取类型',
      icon: 'sum',
      kinds: ['sheet'],
      requires: ['selection.read'],
      phrases: [],
      readOptions: { includeCellTypes: true },
      prepare: () => ({ summary: '读取完整选区', plan: { type: 'report', message: '只读结果' } }),
    };
    const runner = new TaskRunner(host, new FeatureRegistry([detailedReport]));
    const preview = await runner.prepare(detailedReport.id);
    state.valueTypes = [['Error']];

    await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(runner.history).toEqual([]);
    expect(writes).not.toHaveBeenCalled();
  });
});
