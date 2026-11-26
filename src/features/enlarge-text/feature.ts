import type { FeatureDefinition } from '../../core/contracts';
import { HelperError } from '../../core/errors';

export const enlargeText: FeatureDefinition = {
  id: 'text.enlarge',
  title: '把字放大',
  description: '选中的文字变大，看起来更清楚',
  icon: 'text',
  kinds: ['text'],
  requires: ['selection.read', 'text.resize'],
  phrases: ['把字放大', '字体放大', '字大一点', '看不清', '放大文字'],
  prepare(selection) {
    if (!selection.text?.trim())
      throw new HelperError('EMPTY_SELECTION', '请先在文档里选中想放大的文字。');
    if (!selection.fontSize)
      throw new HelperError('UNAVAILABLE', '这段文字包含不同字号，请先选中字号相同的一段。');
    const size = Math.min(72, Math.max(18, (selection.fontSize ?? 14) + 4));
    if (selection.fontSize === size)
      throw new HelperError('INVALID_PLAN', '文字已经足够大了，这次不需要修改。');
    return {
      summary: `将选中的文字调整为 ${size} 磅，文字内容保持不变。`,
      plan: { type: 'text.resize', size },
    };
  },
};
