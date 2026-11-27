import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { Worker } from 'node:worker_threads';
import { transformWithEsbuild } from 'vite';
import type { CellType, CellValue, SelectionSnapshot } from '../../src/core/contracts';
import { FeatureRegistry } from '../../src/core/registry';
import { features } from '../../src/features';
import { averageSelection } from '../../src/features/average-selection/feature';
import { LocalIntentProvider } from '../../src/services/intent';

function selection(values: CellValue[][], cellTypes?: CellType[][]): SelectionSnapshot {
  return {
    documentId: 'average-test',
    documentName: '平均值测试.xlsx',
    kind: 'sheet',
    targetId: 'sheet1!A1:H5',
    targetLabel: 'A1:H5',
    revision: '0',
    values,
    cellTypes:
      cellTypes ??
      values.map((row) =>
        row.map((value) =>
          value === null || value === ''
            ? 'empty'
            : typeof value === 'number'
              ? 'number'
              : typeof value === 'boolean'
                ? 'boolean'
                : 'text',
        ),
      ),
    rowScope: 'all',
  };
}

function message(snapshot: SelectionSnapshot): string {
  const { plan } = averageSelection.prepare(snapshot);
  if (plan.type !== 'report') throw new Error('平均值只能生成只读报告');
  return plan.message;
}

const formatted = (value: number) =>
  new Intl.NumberFormat('zh-CN', { maximumSignificantDigits: 15 }).format(value);

describe('只读平均值的数值与范围规则', () => {
  it.each([
    [[120, 85, 45.5, 200], 112.625],
    [[0.1, 0.2], 0.15],
    [[-6, 0, 9], 1],
    [[0, -0], 0],
    [[Number.MAX_VALUE, Number.MAX_VALUE], Number.MAX_VALUE],
    [[1e16, 1, -1e16], 1 / 3],
    [[Number.MAX_VALUE, -Number.MAX_VALUE, 1e-100], 1e-100 / 3],
    [[Number.MIN_VALUE, Number.MIN_VALUE, Number.MIN_VALUE, 0, 0], Number.MIN_VALUE],
  ])('计算 %j 的平均值且避免中间溢出或抵消丢失', (values, expected) => {
    const result = message(selection([values as number[]]));
    expect(result).toContain(`平均值 ${formatted(expected as number)}，`);
    expect(result).toContain(`共 ${(values as number[]).length} 个数字`);
    expect(result).not.toContain('平均值 -0，');
  });

  it('统计所有行的数据、保留零，并分别报告忽略的文字、空白和布尔值', () => {
    const snapshot = selection([[0, 12, '100', '#DIV/0!', null, '', true, false]]);
    const before = structuredClone(snapshot);
    const result = message(snapshot);
    expect(result).toContain('平均值 6，共 2 个数字');
    expect(result).toContain('文字 2 格、空白 2 格、布尔值 2 格未计入');
    expect(result).toContain('包含选区内的隐藏行和筛选行');
    expect(result).toContain('结果仅显示在这里');
    expect(snapshot).toEqual(before);
    expect(averageSelection.prepare(snapshot).summary).toContain('不写回单元格');
  });

  it('公式返回的空字符串按宿主类型归为文字，不猜测为数值或错误', () => {
    expect(message(selection([[10, '']], [['number', 'text']]))).toContain('文字 1 格、空白 0 格');
  });

  it.each([
    [
      [12, '#VALUE!'],
      ['number', 'error'],
    ],
    [
      [12, 2007],
      ['number', 'error'],
    ],
    [
      ['#N/A', '#REF!'],
      ['error', 'error'],
    ],
  ])('遇到错误单元格直接拒绝，不报告不完整平均值：%j', (values, types) => {
    expect(() =>
      averageSelection.prepare(selection([values as CellValue[]], [types as CellType[]])),
    ).toThrow('错误单元格');
  });

  it.each([NaN, Infinity, -Infinity])('拒绝非有限数字 %s', (value) => {
    expect(() => averageSelection.prepare(selection([[1, value]]))).toThrow('非有限数字');
  });

  it.each([{ values: [] }, { values: [[]] }, { values: [['文字', null, false]] }])(
    '没有数值时不把平均值误报为零：%j',
    ({ values }) => {
      expect(() => averageSelection.prepare(selection(values as CellValue[][]))).toThrow(
        '没有数字',
      );
    },
  );

  it.each([
    { cellTypes: undefined },
    { values: undefined },
    { rowScope: undefined },
    { kind: 'text' as const },
    { cellTypes: [] },
    { cellTypes: [['number'], ['number']] as CellType[][] },
    { cellTypes: [['number', 'unknown']] as CellType[][] },
    { cellTypes: [['number', 'text']] as CellType[][] },
    { cellTypes: [['number', 'empty']] as CellType[][] },
    { cellTypes: [['number', 'boolean']] as CellType[][] },
  ])('拒绝缺失、未知或与值不一致的类型信息：%j', (patch) => {
    expect(() => averageSelection.prepare({ ...selection([[12, 24]]), ...patch })).toThrow(
      '无法可靠读取',
    );
  });

  it('拒绝不完整的二维矩阵', () => {
    expect(() => averageSelection.prepare(selection([[12], [24, 36]]))).toThrow('无法可靠读取');
  });

  it('不丢弃调用者传入的格式参数来执行默认查询', () => {
    expect(() => averageSelection.prepare(selection([[12, 24]]), { fontSize: 30 })).toThrow(
      '不接受目标数值或格式参数',
    );
  });
});

describe('平均值指令只接受完整的单步查询', () => {
  const provider = new LocalIntentProvider();
  it.each([
    '平均值',
    '平均数',
    '求平均',
    '求平均值',
    '求一下平均值',
    '求一下选中数字的平均值',
    '看看一下平均值',
    '算一下平均值',
    '帮我算一下选中范围的平均值',
    '请计算这些数字的平均数。',
    '请帮我算一下当前选中的单元格的平均值，谢谢',
    '选区求平均数',
  ])('接受 %s', async (input) => {
    expect(await provider.resolve(input, features)).toEqual({
      type: 'match',
      featureId: 'sheet.average',
    });
  });

  it.each([
    '不要算平均值',
    '平均值并删除单元格',
    '算平均值再写入C1',
    '只计算可见行的平均值',
    '排除隐藏行求平均',
    '求筛选后可见数据的平均值',
    '平均值设为30',
    '平均值保留2位小数',
    '平均数写回表格',
    '我昨天听说了平均值',
    '求平均值完成后保存',
    '平均值+30',
  ])('要求澄清 %s', async (input) => {
    expect(await provider.resolve(input, features)).toMatchObject({ type: 'clarify' });
  });

  it('不默认选择合计与平均值混合请求中的一项', async () => {
    expect(await provider.resolve('先求和，再求平均值', features)).toEqual({
      type: 'choices',
      featureIds: ['sheet.sum', 'sheet.average'],
    });
  });

  it('功能关闭后指令入口同时不可用', async () => {
    const registry = new FeatureRegistry(features, { 'sheet.average': false });
    expect(await provider.resolve('求平均值', registry.list())).toEqual({ type: 'unknown' });
  });

  it('重复目标和长无效请求有界拒绝，不产生指数回溯', async () => {
    // Isolate the actual parser in a worker so a backtracking regression can be
    // terminated instead of hanging the test runner's own JavaScript thread.
    const source = await readFile(new URL('../../src/services/intent.ts', import.meta.url), 'utf8');
    const compiled = await transformWithEsbuild(source, 'intent.ts', {
      loader: 'ts',
      target: 'es2022',
    });
    const worker = new Worker(
      `const { parentPort, workerData } = require('node:worker_threads');
         import(workerData.moduleUrl).then(async ({ LocalIntentProvider }) => {
           const provider = new LocalIntentProvider();
           const result = [];
           for (const input of workerData.inputs)
             result.push(await provider.resolve(input, [{ id: 'sheet.average', phrases: ['平均值'] }]));
           parentPort.postMessage(result);
         }).catch(error => { throw error; });`,
      {
        eval: true,
        workerData: {
          moduleUrl: `data:text/javascript;base64,${Buffer.from(compiled.code).toString('base64')}`,
          inputs: [
            `${'选中数字'.repeat(30)}平均值x`,
            `${'选中数字'.repeat(10000)}平均值x`,
            `${'这些数字的'.repeat(10000)}平均值x`,
            `${'请帮我'.repeat(10000)}平均值x`,
          ],
        },
      },
    );
    try {
      const results = await new Promise<unknown[]>((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('平均值指令解析超过5秒')), 5000);
        worker.once('message', (result: unknown[]) => {
          clearTimeout(timer);
          resolve(result);
        });
        worker.once('error', (error) => {
          clearTimeout(timer);
          reject(error);
        });
      });
      expect(results).toHaveLength(4);
      results.forEach((result) => expect(result).toMatchObject({ type: 'clarify' }));
    } finally {
      await worker.terminate();
    }
  }, 10000);
});
