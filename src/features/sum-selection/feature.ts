import type { FeatureDefinition } from '../../core/contracts';
import { HelperError } from '../../core/errors';

export const sumSelection: FeatureDefinition = {
  id: 'sheet.sum',
  title: '算一下合计',
  description: '把选中范围里的数字加起来',
  icon: 'sum',
  kinds: ['sheet'],
  requires: ['selection.read'],
  phrases: ['算一下合计', '合计', '求和', '总共多少钱', '算总数', '加起来'],
  prepare(selection) {
    const cells = selection.values?.flat() ?? [];
    const numbers = cells.filter(
      (value): value is number => typeof value === 'number' && Number.isFinite(value),
    );
    if (!numbers.length)
      throw new HelperError('EMPTY_SELECTION', '选中范围里没有数字。请先选中要相加的数字。');
    // Kahan 求和减少累计误差；显示精度有限，不替代财务计算引擎。
    let total = 0;
    let compensation = 0;
    for (const value of numbers) {
      const adjusted = value - compensation;
      const next = total + adjusted;
      compensation = next - total - adjusted;
      total = next;
    }
    if (!Number.isFinite(total))
      throw new HelperError('INVALID_PLAN', '数字合计超出了可处理范围，请缩小范围后再试。');
    const formatted = new Intl.NumberFormat('zh-CN', { maximumSignificantDigits: 15 }).format(
      total,
    );
    const ignored = cells.length - numbers.length;
    const message = `合计 ${formatted}，共 ${numbers.length} 个数字${ignored ? `；另有 ${ignored} 个文字、空白或非数字单元格未计入` : ''}。结果仅显示在这里。`;
    return {
      summary: `计算「${selection.targetLabel}」中的数字，在助手中显示合计。`,
      plan: { type: 'report', message },
    };
  },
};
