import { describe, expect, it, vi } from 'vitest';
import type { CellValue } from '../../core/contracts';
import { FeatureRegistry } from '../../core/registry';
import { TaskRunner } from '../../core/runner';
import { averageSelection } from '../../features/average-selection/feature';
import { WpsAdapter } from './adapter';
import type { WpsApplication, WpsSelection, WpsSheetCell } from './bridge';

function fixture(initial: CellValue[][] = [[120], [85], [45.5], [200]]) {
  const state = { values: structuredClone(initial), areas: 1 };
  const positions = new WeakMap<WpsSheetCell, [number, number]>();
  const errors = new Set<string>();
  const item = vi.fn((row: number, column: number) => {
    const cell: WpsSheetCell = {
      get Value2() {
        return state.values[row - 1]?.[column - 1];
      },
    };
    positions.set(cell, [row - 1, column - 1]);
    return cell;
  });
  const isError = vi.fn((cell: WpsSheetCell): boolean | number => {
    const position = positions.get(cell);
    if (!position) throw new Error('Must receive a native cell reference');
    return errors.has(position.join(':'));
  });
  const readValues = vi.fn(() =>
    initial.length === 1 && initial[0]?.length === 1
      ? state.values[0]![0]!
      : structuredClone(state.values),
  );
  const selection: WpsSelection = {
    Rows: { Count: initial.length },
    Columns: { Count: initial[0]?.length ?? 0 },
    Address: '$B$2:$B$5',
    get Value2() {
      return readValues();
    },
    Areas: {
      get Count() {
        return state.areas;
      },
    },
    Cells: { Item: item },
  };
  const app: WpsApplication = {
    ActiveWorkbook: { Name: '费用.xlsx' },
    ActiveSheet: { Name: '费用' },
    Selection: selection,
    WorksheetFunction: { IsError: isError },
  };
  const host = new WpsAdapter('sheet', () => app);
  return { host, app, selection, state, errors, item, isError, readValues };
}

describe('WPS 按需读取平均值类型', () => {
  it('普通读取不访问类型接口，保留现有合计路径', async () => {
    const { host, selection, item, isError } = fixture();
    Object.defineProperty(selection, 'Areas', {
      get() {
        throw new Error('Legacy reading must not probe Areas');
      },
    });
    const result = await host.readSelection();
    expect(result.values).toEqual([[120], [85], [45.5], [200]]);
    expect(result.cellTypes).toBeUndefined();
    expect(result.rowScope).toBeUndefined();
    expect(item).not.toHaveBeenCalled();
    expect(isError).not.toHaveBeenCalled();
  });

  it('逐个原生引用判错，并读取完整连续范围且不写入', async () => {
    const { host, state, item, isError, readValues } = fixture([
      [0, '100', null],
      [false, '#VALUE!', 42],
    ]);
    const before = structuredClone(state);
    const result = await host.readSelection({ includeCellTypes: true });
    expect(result).toMatchObject({
      values: before.values,
      cellTypes: [
        ['number', 'text', 'empty'],
        ['boolean', 'text', 'number'],
      ],
      rowScope: 'all',
    });
    expect(item.mock.calls).toEqual([
      [1, 1],
      [1, 2],
      [1, 3],
      [2, 1],
      [2, 2],
      [2, 3],
    ]);
    expect(isError).toHaveBeenCalledTimes(6);
    expect(readValues).toHaveBeenCalledTimes(2);
    expect(state).toEqual(before);
  });

  it.each([true, -1])('原生错误结果 %s 能识别编码成数字的错误值', async (error) => {
    const { host, isError } = fixture([[2042]]);
    isError.mockReturnValue(error);
    expect(await host.readSelection({ includeCellTypes: true })).toMatchObject({
      values: [[2042]],
      cellTypes: [['error']],
    });
  });

  it.each([false, 0])('原生非错误结果 %s 保持数字类型', async (error) => {
    const { host, isError } = fixture([[2042]]);
    isError.mockReturnValue(error);
    expect(await host.readSelection({ includeCellTypes: true })).toMatchObject({
      cellTypes: [['number']],
    });
  });

  it('实际字符串错误与普通错误样式文本依据原生结果区分', async () => {
    const { host, errors } = fixture([['#N/A', '#N/A']]);
    errors.add('0:0');
    expect(await host.readSelection({ includeCellTypes: true })).toMatchObject({
      cellTypes: [['error', 'text']],
    });
  });

  it('平均值查询在确认后只报告112.625，不修改单元格且不能恢复', async () => {
    const { host, state } = fixture();
    const before = structuredClone(state);
    const apply = vi.spyOn(host, 'apply');
    const runner = new TaskRunner(host, new FeatureRegistry([averageSelection]));
    const preview = await runner.prepare('sheet.average');
    expect(preview.changesDocument).toBe(false);
    const result = await runner.execute(preview.id);
    expect(result).toMatchObject({ canUndo: false });
    expect(result.message).toContain('112.625');
    expect(apply).not.toHaveBeenCalled();
    expect(state).toEqual(before);
  });

  it('值未改变但原生类型变为错误时，确认拒绝旧预览', async () => {
    const { host, errors } = fixture([[2042]]);
    const runner = new TaskRunner(host, new FeatureRegistry([averageSelection]));
    const preview = await runner.prepare('sheet.average');
    errors.add('0:0');
    await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(runner.history).toEqual([]);
  });

  it('缺少原生类型接口时平均值明确不可用，但合计原路径仍可读', async () => {
    const { host, app } = fixture();
    delete app.WorksheetFunction;
    const runner = new TaskRunner(host, new FeatureRegistry([averageSelection]));
    await expect(runner.prepare('sheet.average')).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    expect((await host.readSelection()).values).toEqual([[120], [85], [45.5], [200]]);
  });

  it.each(['Areas', 'Cells', 'WorksheetFunction'] as const)(
    '缺少 %s 时不伪造可靠类型',
    async (key) => {
      const { host, app, selection } = fixture();
      if (key === 'WorksheetFunction') delete app.WorksheetFunction;
      else delete selection[key];
      const result = await host.readSelection({ includeCellTypes: true });
      expect(result.cellTypes).toBeUndefined();
      expect(result.rowScope).toBeUndefined();
      expect(result.values).toEqual([[120], [85], [45.5], [200]]);
    },
  );

  it('多区域选区不宣称已读取全部选中行', async () => {
    const { host, state, item } = fixture();
    state.areas = 2;
    const result = await host.readSelection({ includeCellTypes: true });
    expect(result.cellTypes).toBeUndefined();
    expect(result.rowScope).toBeUndefined();
    expect(item).not.toHaveBeenCalled();
  });

  it.each([1, 2, NaN])('未知判错返回值 %s 不被猜测成布尔值', async (unknown) => {
    const { host, isError } = fixture();
    isError.mockReturnValue(unknown);
    expect((await host.readSelection({ includeCellTypes: true })).cellTypes).toBeUndefined();
  });

  it('原生判错接口抛异常时不会使平均值得到可用类型', async () => {
    const { host, isError } = fixture();
    isError.mockImplementation(() => {
      throw new Error('Unsupported function');
    });
    expect((await host.readSelection({ includeCellTypes: true })).cellTypes).toBeUndefined();
  });

  it('逐格值与批量快照不一致时不提交类型', async () => {
    const { host, item } = fixture();
    item.mockReturnValue({ Value2: 999 });
    expect((await host.readSelection({ includeCellTypes: true })).cellTypes).toBeUndefined();
  });

  it('类型探测后整体值改变时不提交过时类型', async () => {
    const { host, readValues } = fixture();
    readValues.mockReturnValueOnce([[120], [85], [45.5], [200]]).mockReturnValueOnce([[999]]);
    const result = await host.readSelection({ includeCellTypes: true });
    expect(result.cellTypes).toBeUndefined();
    expect(result.values).toEqual([[120], [85], [45.5], [200]]);
  });

  it('非有限数值被标为未知，不能借typeof number冒充有效数字', async () => {
    const { host } = fixture([[Infinity]]);
    expect(await host.readSelection({ includeCellTypes: true })).toMatchObject({
      cellTypes: [['unknown']],
    });
  });

  it('超限范围在读取数据或逐格探测之前拒绝', async () => {
    const { host, selection, readValues, item, isError } = fixture();
    selection.Rows = { Count: 5001 };
    await expect(host.readSelection({ includeCellTypes: true })).rejects.toMatchObject({
      code: 'SELECTION_TOO_LARGE',
    });
    expect(readValues).not.toHaveBeenCalled();
    expect(item).not.toHaveBeenCalled();
    expect(isError).not.toHaveBeenCalled();
  });
});
