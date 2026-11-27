import { OfficeAdapter } from '../adapters/office/adapter';
import { WpsAdapter } from '../adapters/wps/adapter';
import { getWpsApplication } from '../adapters/wps/bridge';
import { createRuntime } from '../app/runtime';
import type { CellValue, HostAdapter, SelectionSnapshot } from '../core/contracts';
import { assertSameSelection } from '../core/selection';

const fixtureName = 'OfficeHelper-Native-Test.xlsx';
const fixtureMarker = 'OFFICE_HELPER_NATIVE_TEST';
const fixtureValues = [[120], [85], [45.5], [200]];
const detailOptions = { includeCellTypes: true };

interface NativeFixture {
  sheet: string;
  values: CellValue[][];
}

function assertFixture(selection: SelectionSnapshot) {
  const address = selection.targetLabel.split('!').at(-1) ?? '';
  if (
    selection.kind !== 'sheet' ||
    selection.documentName !== fixtureName ||
    !selection.documentId ||
    !selection.targetId ||
    !/^\$?B\$?2:\$?B\$?5$/.test(address) ||
    JSON.stringify(selection.values) !== JSON.stringify(fixtureValues)
  )
    throw new Error('平均值验收只允许专用测试文件的 B2:B5 四个固定数值，未执行查询。');
  if (
    selection.rowScope !== 'all' ||
    JSON.stringify(selection.cellTypes) !== '[["number"],["number"],["number"],["number"]]'
  )
    throw new Error('平均值验收无法可靠读取完整范围及四个数字的原生类型，停止验收。');
}

function fixtureMatrix(raw: unknown): CellValue[][] {
  if (
    !Array.isArray(raw) ||
    raw.length !== 5 ||
    raw.some(
      (row) =>
        !Array.isArray(row) ||
        row.length !== 2 ||
        row.some(
          (value) =>
            value !== null &&
            typeof value !== 'string' &&
            typeof value !== 'boolean' &&
            !(typeof value === 'number' && Number.isFinite(value)),
        ),
    ) ||
    raw[0]?.[0] !== fixtureMarker ||
    JSON.stringify(raw.slice(1).map((row: CellValue[]) => [row[1]])) !==
      JSON.stringify(fixtureValues)
  )
    throw new Error('专用工作表 A1 标记或 A1:B5 样本数据不匹配，停止验收。');
  return raw.map((row: CellValue[]) => [...row]);
}

async function readNativeFixture(adapter: HostAdapter): Promise<NativeFixture> {
  if (adapter instanceof OfficeAdapter && adapter.info.id === 'office') {
    if (
      typeof Office === 'undefined' ||
      typeof Excel === 'undefined' ||
      Office.context.host !== Office.HostType.Excel
    )
      throw new Error('平均值验收必须在真实 Excel 任务窗格运行。');
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const fixture = sheet.getRange('A1:B5');
      sheet.load('id');
      fixture.load('values');
      await context.sync();
      return { sheet: sheet.id, values: fixtureMatrix(fixture.values) };
    });
  }
  if (adapter instanceof WpsAdapter && adapter.info.id === 'wps') {
    const app = getWpsApplication('sheet');
    const sheet = app.ActiveSheet as
      { Name: string; Range(address: string): { Value2: unknown } } | undefined;
    if (app.ActiveWorkbook?.Name !== fixtureName || !sheet || typeof sheet.Range !== 'function')
      throw new Error('平均值验收必须在真实 WPS 表格任务窗格与专用样本运行。');
    return { sheet: sheet.Name, values: fixtureMatrix(sheet.Range('A1:B5').Value2) };
  }
  throw new Error('平均值验收不接受演示或自定义适配器，必须使用 Office/WPS 原生表格连接。');
}

/** Token-gated development caller only. Every operation is a read or report. */
export async function checkSheetAverage(
  adapter: HostAdapter,
  report: (stage: string, details?: Record<string, unknown>) => Promise<void>,
) {
  if (
    !import.meta.env.DEV ||
    adapter.info.kind !== 'sheet' ||
    !(
      (adapter instanceof OfficeAdapter && adapter.info.id === 'office') ||
      (adapter instanceof WpsAdapter && adapter.info.id === 'wps')
    )
  )
    throw new Error('平均值专项只允许开发环境中的真实 Office/WPS 表格适配器。');

  const readGuarded = async () => {
    const before = await adapter.readSelection(detailOptions);
    assertFixture(before);
    const native = await readNativeFixture(adapter);
    const after = await adapter.readSelection(detailOptions);
    assertFixture(after);
    assertSameSelection(before, after);
    return { selection: after, native };
  };
  const original = await readGuarded();
  const unchanged = async () => {
    const current = await readGuarded();
    assertSameSelection(original.selection, current.selection);
    if (JSON.stringify(original.native) !== JSON.stringify(current.native))
      throw new Error('平均值验收期间工作表标记、样本值或工作表归属发生变化。');
  };
  const readonlyStep = async <T>(operation: () => Promise<T>): Promise<T> => {
    await unchanged();
    try {
      return await operation();
    } finally {
      await unchanged();
    }
  };
  const { runner } = createRuntime(adapter);
  await report('sheet-average-fixture-verified', { cells: 4, readonly: true });

  let preview = await readonlyStep(() => runner.prepare('sheet.average'));
  if (preview.changesDocument || !preview.summary.includes('隐藏行和筛选行'))
    throw new Error('平均值预览没有明确只读与完整行范围。');
  await readonlyStep(async () => runner.cancel(preview.id));
  let cancelled = false;
  await readonlyStep(async () => {
    try {
      await runner.execute(preview.id);
    } catch (error) {
      if ((error as { code?: string }).code !== 'EXPIRED_TASK') throw error;
      cancelled = true;
    }
  });
  if (!cancelled) throw new Error('已取消的平均值预览仍然可以执行。');
  await report('sheet-average-cancel-verified', { readonly: true });

  preview = await readonlyStep(() => runner.prepare('sheet.average'));
  if (preview.changesDocument) throw new Error('平均值查询被错误标记为文档修改。');
  const average = await readonlyStep(() => runner.execute(preview.id));
  if (!average.message.startsWith('平均值 112.625，') || average.canUndo || runner.canUndo)
    throw new Error('平均值结果或只读恢复状态不正确。');
  await report('sheet-average-result-verified', { average: 112.625, readonly: true });

  const sumPreview = await readonlyStep(() => runner.prepare('sheet.sum'));
  if (sumPreview.changesDocument) throw new Error('原有合计被错误标记为文档修改。');
  const sum = await readonlyStep(() => runner.execute(sumPreview.id));
  if (!sum.message.startsWith('合计 450.5，') || sum.canUndo || runner.canUndo)
    throw new Error('原有合计结果或只读恢复状态发生变化。');
  await unchanged();
  await report('sheet-average-passed', {
    average: 112.625,
    total: 450.5,
    cells: 4,
    cancelled: true,
    canUndo: false,
    readonly: true,
  });
}
