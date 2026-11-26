import type { FeatureDefinition } from '../../core/contracts';

export const inspectSelection: FeatureDefinition = {
  id: 'selection.inspect',
  title: '看看选中了什么',
  description: '先确认位置，再放心操作',
  icon: 'inspect',
  kinds: ['text', 'sheet'],
  requires: ['selection.read'],
  phrases: ['看看选中了什么', '选中内容', '检查范围', '看看内容'],
  prepare(selection) {
    const message =
      selection.kind === 'text'
        ? `当前选中 ${Array.from(selection.text ?? '').length} 个字符${selection.fontSize ? `，字号 ${selection.fontSize} 磅` : ''}。`
        : `当前选中 ${selection.values?.length ?? 0} 行，共 ${selection.values?.reduce((n, row) => n + row.length, 0) ?? 0} 个单元格。`;
    return {
      summary: `查看「${selection.targetLabel}」的信息。`,
      plan: { type: 'report', message },
    };
  },
};
