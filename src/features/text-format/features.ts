import type { FeatureDefinition, SelectionSnapshot, TextActionPlan } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { projectTextPlan, requireParagraphs, validateTextSelection } from '../../core/text-format';

const alignmentNames: Record<string, string> = {
  Left: '左对齐',
  Centered: '居中',
  Right: '右对齐',
  Justified: '两端对齐',
};

function prepareParagraph(selection: SelectionSnapshot, plan: TextActionPlan) {
  validateTextSelection(plan, selection);
  const before = requireParagraphs(selection);
  const after = projectTextPlan(selection, plan).paragraphs!;
  if (JSON.stringify(before) === JSON.stringify(after))
    throw new HelperError('INVALID_PLAN', '这些段落已经是要求的格式，这次不需要修改。');
  const changes = before.map((p, index) => {
    const next = after[index]!;
    const snippet = p.text.replace(/[\r\n]/g, '').slice(0, 24);
    const values =
      plan.type === 'paragraph.center'
        ? `${alignmentNames[p.alignment]} → 居中`
        : plan.type === 'paragraph.spacing'
          ? `${p.lineSpacing / 12} 倍 → ${next.lineSpacing / 12} 倍行距`
          : `首行缩进 ${Number(p.firstLineIndent.toFixed(2))} → ${Number(next.firstLineIndent.toFixed(2))} 磅`;
    return `第 ${index + 1} 段“${snippet}${p.text.length > 24 ? '…' : ''}”：${values}。`;
  });
  return {
    summary: `将处理选中文字所在的 ${before.length} 个完整段落（包括未选中的同段文字）。${changes.join('')}格式修改可能改变分页。`,
    plan,
  };
}

export const boldText: FeatureDefinition = {
  id: 'text.bold',
  title: '加粗文字',
  description: '加粗选中的文字，也可输入“取消加粗”',
  icon: 'text',
  kinds: ['text'],
  requires: ['selection.read', 'text.bold'],
  phrases: ['加粗', '粗体'],
  prepare(selection, parameters) {
    const plan: TextActionPlan = { type: 'text.bold', bold: parameters?.bold ?? true };
    validateTextSelection(plan, selection);
    if (selection.bold === plan.bold)
      throw new HelperError('INVALID_PLAN', '选中文字已经是要求的粗细，这次不需要修改。');
    return {
      summary: `只处理选中的文字：${selection.bold ? '加粗' : '常规'} → ${plan.bold ? '加粗' : '常规'}。文字内容保持不变，分页可能改变。`,
      plan,
    };
  },
};

export const centerParagraph: FeatureDefinition = {
  id: 'paragraph.center',
  title: '段落居中',
  description: '将选中文字所在的完整段落居中',
  icon: 'text',
  kinds: ['text'],
  requires: ['selection.read', 'paragraph.center'],
  phrases: ['段落居中', '文字居中', '居中对齐', '居中'],
  prepare: (selection) => prepareParagraph(selection, { type: 'paragraph.center' }),
};

export const spaceParagraph: FeatureDefinition = {
  id: 'paragraph.spacing',
  title: '设置行距',
  description: '输入具体倍数，例如“行距设为1.5倍”',
  icon: 'text',
  kinds: ['text'],
  requires: ['selection.read', 'paragraph.spacing'],
  phrases: ['行距'],
  inputPrompt: {
    prefix: '行距设为',
    message: '请在上方补上你想要的倍数，例如“行距设为1.5倍”（支持1–3倍）。',
  },
  prepare(selection, parameters) {
    if (parameters?.lineSpacingMultiple === undefined)
      throw new HelperError(
        'INVALID_PLAN',
        '请在上方输入具体行距，例如“行距设为1.5倍”（支持1–3倍）。',
      );
    return prepareParagraph(selection, {
      type: 'paragraph.spacing',
      multiple: parameters.lineSpacingMultiple,
    });
  },
};

export const indentParagraph: FeatureDefinition = {
  id: 'paragraph.indent',
  title: '设置首行缩进',
  description: '输入明确数值和单位，例如“首行缩进24磅”',
  icon: 'text',
  kinds: ['text'],
  requires: ['selection.read', 'paragraph.indent'],
  phrases: ['首行缩进', '缩进'],
  inputPrompt: {
    prefix: '首行缩进',
    message: '请在上方补上数值和单位，例如“首行缩进24磅”或“首行缩进1厘米”。',
  },
  prepare(selection, parameters) {
    if (parameters?.firstLineIndentPoints === undefined)
      throw new HelperError(
        'INVALID_PLAN',
        '请在上方输入明确的首行缩进，例如“首行缩进24磅”或“首行缩进1厘米”。',
      );
    return prepareParagraph(selection, {
      type: 'paragraph.indent',
      points: parameters.firstLineIndentPoints,
    });
  },
};
