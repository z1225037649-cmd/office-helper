import type { ActionPlan, ParagraphSnapshot, SelectionSnapshot, TextActionPlan } from './contracts';
import { HelperError } from './errors';

export function validateTextPlan(plan: ActionPlan): asserts plan is TextActionPlan {
  switch (plan.type) {
    case 'text.resize':
      if (
        Number.isFinite(plan.size) &&
        plan.size >= 8 &&
        plan.size <= 72 &&
        Number.isInteger(plan.size * 2)
      )
        return;
      break;
    case 'text.bold':
      if (typeof plan.bold === 'boolean') return;
      break;
    case 'text.documentStyles': {
      if (
        Array.isArray(plan.assignments) &&
        plan.assignments.length > 0 &&
        plan.assignments.length <= 1500 &&
        plan.assignments.every(
          (assignment) =>
            assignment &&
            Number.isInteger(assignment.paragraphIndex) &&
            assignment.paragraphIndex >= 0 &&
            ['Heading1', 'Heading2', 'Heading3'].includes(assignment.style),
        ) &&
        new Set(plan.assignments.map((assignment) => assignment.paragraphIndex)).size ===
          plan.assignments.length
      )
        return;
      break;
    }
    case 'paragraph.center':
      return;
    case 'paragraph.spacing':
      if (
        Number.isFinite(plan.multiple) &&
        plan.multiple >= 1 &&
        plan.multiple <= 3 &&
        Math.abs(plan.multiple * 20 - Math.round(plan.multiple * 20)) < 1e-8
      )
        return;
      break;
    case 'paragraph.indent':
      if (
        Number.isFinite(plan.points) &&
        plan.points >= 0 &&
        plan.points <= 144 &&
        Math.abs(plan.points * 20 - Math.round(plan.points * 20)) < 1e-8
      )
        return;
      break;
  }
  throw new HelperError('INVALID_PLAN', '这个格式参数暂不支持，请重新说明具体要求。');
}

export function requireParagraphs(selection: SelectionSnapshot): ParagraphSnapshot[] {
  if (!selection.text?.trim()) throw new HelperError('EMPTY_SELECTION', '请先选中要处理的文字。');
  if (!selection.paragraphs?.length)
    throw new HelperError('UNAVAILABLE', '当前版本不能准确读取段落格式，暂时无法修改。');
  return selection.paragraphs;
}

export function validateTextSelection(plan: TextActionPlan, selection: SelectionSnapshot): void {
  validateTextPlan(plan);
  if (plan.type === 'text.documentStyles') {
    const paragraphs = selection.documentParagraphs;
    if (selection.kind !== 'text' || selection.targetId !== 'document-body' || !paragraphs?.length)
      throw new HelperError('UNAVAILABLE', '无法读取全文段落，请重新打开文档后再试。');
    if (
      plan.assignments.some(({ paragraphIndex }) => {
        const paragraph = paragraphs[paragraphIndex];
        return (
          !paragraph ||
          !paragraph.text.trim() ||
          paragraph.isListItem ||
          paragraph.inTable ||
          paragraph.style !== 'Normal'
        );
      })
    )
      throw new HelperError(
        'STALE_SELECTION',
        '需要处理的段落已改变或含列表、表格、自定义样式，请重新检查全文。',
      );
    return;
  }
  if (selection.kind !== 'text' || !selection.text?.trim())
    throw new HelperError('EMPTY_SELECTION', '请先选中要处理的文字。');
  if (plan.type === 'text.resize' && !selection.fontSize)
    throw new HelperError('UNAVAILABLE', '请先选中字号相同的一段文字，便于准确恢复。');
  if (plan.type === 'text.bold' && typeof selection.bold !== 'boolean')
    throw new HelperError('UNAVAILABLE', '选中文字的粗细不一致，请先选择粗细相同的一段。');
  if (!plan.type.startsWith('paragraph.')) return;
  const paragraphs = requireParagraphs(selection);
  if (
    plan.type === 'paragraph.center' &&
    paragraphs.some((p) => !['Left', 'Centered', 'Right', 'Justified'].includes(p.alignment))
  )
    throw new HelperError('UNAVAILABLE', '这段文字的对齐方式暂不支持准确恢复。');
  if (
    plan.type === 'paragraph.spacing' &&
    paragraphs.some(
      (p) => p.spacingMode !== 'multiple' || !Number.isFinite(p.lineSpacing) || p.lineSpacing <= 0,
    )
  )
    throw new HelperError(
      'UNAVAILABLE',
      '当前行距无法准确恢复，请先在文档中设置明确的多倍行距，再重新预览。',
    );
  if (
    plan.type === 'paragraph.indent' &&
    paragraphs.some(
      (p) =>
        p.isListItem ||
        !Number.isFinite(p.firstLineIndent) ||
        p.firstLineIndentRestorable === false,
    )
  )
    throw new HelperError(
      'UNAVAILABLE',
      '当前缩进无法准确恢复，请选择非列表、可读取磅值缩进的普通段落。',
    );
}

/** Expected effective state; no native objects or writes. */
export function projectTextPlan(
  selection: SelectionSnapshot,
  plan: TextActionPlan,
): SelectionSnapshot {
  const next = structuredClone(selection);
  if (plan.type === 'text.resize') next.fontSize = plan.size;
  else if (plan.type === 'text.bold') next.bold = plan.bold;
  else if (plan.type === 'text.documentStyles')
    for (const assignment of plan.assignments) {
      const paragraph = next.documentParagraphs?.[assignment.paragraphIndex];
      if (paragraph) paragraph.style = assignment.style;
    }
  else
    for (const paragraph of next.paragraphs ?? []) {
      if (plan.type === 'paragraph.center') paragraph.alignment = 'Centered';
      else if (plan.type === 'paragraph.spacing') paragraph.lineSpacing = plan.multiple * 12;
      else if (plan.type === 'paragraph.indent') paragraph.firstLineIndent = plan.points;
    }
  return next;
}

export function textPlanMessage(plan: TextActionPlan): string {
  switch (plan.type) {
    case 'text.resize':
      return `已把选中文字调整为 ${plan.size} 磅。`;
    case 'text.bold':
      return plan.bold ? '已把选中文字加粗。' : '已取消选中文字的加粗。';
    case 'text.documentStyles':
      return `已统一全文 ${plan.assignments.length} 个标题段落的内置样式。`;
    case 'paragraph.center':
      return '已将选中文字所在的完整段落居中。';
    case 'paragraph.spacing':
      return `已将选中文字所在的完整段落设为 ${plan.multiple} 倍行距。`;
    case 'paragraph.indent':
      return `已将选中文字所在的完整段落首行缩进设为 ${Number(plan.points.toFixed(2))} 磅。`;
  }
}
