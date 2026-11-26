import type { ActionPlan, SelectionSnapshot } from '../../core/contracts';
import { HelperError } from '../../core/errors';

export function validateResize(
  plan: ActionPlan,
  selection: SelectionSnapshot,
): asserts plan is { type: 'text.resize'; size: number } {
  if (plan.type !== 'text.resize' || !Number.isFinite(plan.size) || plan.size < 8 || plan.size > 72)
    throw new HelperError('INVALID_PLAN', '字号需要在 8 到 72 磅之间。');
  if (!selection.text?.trim()) throw new HelperError('EMPTY_SELECTION', '请先选中想放大的文字。');
  if (!selection.fontSize || selection.fontSize <= 0 || selection.fontSize > 1638)
    throw new HelperError(
      'UNAVAILABLE',
      '这段文字包含不同字号。请先选中字号相同的一段，便于准确恢复。',
    );
}
