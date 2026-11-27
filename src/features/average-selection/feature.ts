import type { CellType, CellValue, FeatureDefinition } from '../../core/contracts';
import { HelperError } from '../../core/errors';

function unavailable(): never {
  throw new HelperError(
    'UNAVAILABLE',
    '当前选区无法可靠读取每个单元格的类型或完整行范围，暂时不能计算平均值。请重新选择连续范围，或使用支持该读取接口的版本。',
  );
}

function matchesType(value: CellValue, type: CellType): boolean {
  switch (type) {
    case 'number':
      return typeof value === 'number';
    case 'text':
      return typeof value === 'string';
    case 'boolean':
      return typeof value === 'boolean';
    case 'empty':
      return value === null || value === '';
    case 'error':
      return true;
    default:
      return false;
  }
}

/** Accumulate exact binary units, then round once; a finite mean must not overflow mid-sum. */
function average(numbers: readonly number[]): number {
  const view = new DataView(new ArrayBuffer(8));
  const fractionMask = (1n << 52n) - 1n;
  let units = 0n;
  for (const value of numbers) {
    view.setFloat64(0, value);
    const bits = view.getBigUint64(0);
    const exponent = Number((bits >> 52n) & 0x7ffn);
    const fraction = bits & fractionMask;
    const significand = exponent ? fraction | (1n << 52n) : fraction;
    const contribution = significand << BigInt(Math.max(0, exponent - 1));
    units += bits >> 63n ? -contribution : contribution;
  }
  if (units === 0n) return 0;
  const negative = units < 0n;
  const magnitude = negative ? -units : units;
  const count = BigInt(numbers.length);
  // Each unit is 2^-1074. Retain at most 53 significant bits for the final double.
  const shift = Math.max(0, (magnitude / count).toString(2).length - 53);
  const divisor = count << BigInt(shift);
  let significand = magnitude / divisor;
  const remainder = magnitude % divisor;
  if (remainder * 2n > divisor || (remainder * 2n === divisor && significand % 2n !== 0n))
    significand++;
  const result = Number(significand) * 2 ** (shift - 1074);
  return result === 0 ? 0 : negative ? -result : result;
}

export const averageSelection: FeatureDefinition = {
  id: 'sheet.average',
  title: '算一下平均值',
  description: '查看选中数字的平均值，包含隐藏和筛选行',
  icon: 'sum',
  kinds: ['sheet'],
  requires: ['selection.read'],
  readOptions: { includeCellTypes: true },
  phrases: ['平均值', '平均数', '求平均'],
  prepare(selection, parameters) {
    if (parameters && Object.keys(parameters).length)
      throw new HelperError('INVALID_PLAN', '平均值查询不接受目标数值或格式参数，请重新选择功能。');
    const { values, cellTypes } = selection;
    if (
      selection.kind !== 'sheet' ||
      selection.rowScope !== 'all' ||
      !Array.isArray(values) ||
      !Array.isArray(cellTypes) ||
      cellTypes.length !== values.length
    )
      unavailable();

    const numbers: number[] = [];
    const ignored = { text: 0, empty: 0, boolean: 0 };
    let errors = 0;
    const columns = values[0]?.length ?? 0;
    for (let row = 0; row < values.length; row++) {
      const rowValues = values[row]!;
      const rowTypes = cellTypes[row];
      if (
        !Array.isArray(rowValues) ||
        !Array.isArray(rowTypes) ||
        rowValues.length !== columns ||
        rowTypes.length !== rowValues.length
      )
        unavailable();
      for (let column = 0; column < rowValues.length; column++) {
        const value = rowValues[column]!;
        const type = rowTypes[column]!;
        if (!matchesType(value, type)) unavailable();
        if (type === 'error') errors++;
        else if (type === 'number') {
          if (!Number.isFinite(value))
            throw new HelperError('INVALID_PLAN', '选中范围含无法处理的非有限数字，平均值未计算。');
          numbers.push(value as number);
        } else if (type === 'text' || type === 'empty' || type === 'boolean') ignored[type]++;
      }
    }
    if (errors)
      throw new HelperError(
        'INVALID_PLAN',
        `选中范围含 ${errors} 个错误单元格，平均值未计算。请先检查错误值后重新预览。`,
      );
    if (!numbers.length)
      throw new HelperError('EMPTY_SELECTION', '选中范围里没有数字。请先选中要计算平均值的数字。');

    const result = average(numbers);
    if (!Number.isFinite(result))
      throw new HelperError('INVALID_PLAN', '平均值超出了可处理范围，请缩小范围后再试。');
    const formatted = new Intl.NumberFormat('zh-CN', { maximumSignificantDigits: 15 }).format(
      result,
    );
    const exclusions = `文字 ${ignored.text} 格、空白 ${ignored.empty} 格、布尔值 ${ignored.boolean} 格未计入`;
    const scope = '包含选区内的隐藏行和筛选行';
    return {
      summary: `计算「${selection.targetLabel}」中 ${numbers.length} 个数字的平均值；${exclusions}。${scope}，结果只显示在助手中，不写回单元格。`,
      plan: {
        type: 'report',
        message: `平均值 ${formatted}，共 ${numbers.length} 个数字；${exclusions}。${scope}。结果仅显示在这里。`,
      },
    };
  },
};
