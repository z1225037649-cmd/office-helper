import type { FeatureDefinition, TextActionPlan } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { analyzeDocumentStyles } from './analyze';

const headingNames = { Heading1: '一级标题', Heading2: '二级标题', Heading3: '三级标题' } as const;

export const normalizeDocumentStyles: FeatureDefinition = {
  id: 'text.documentStyles',
  title: '统一全文标题样式',
  description: '检查全文的编号标题，只调整能明确识别的标题层级，保留正文、列表和表格',
  icon: 'text',
  kinds: ['text'],
  requires: ['text.documentStyles'],
  readOptions: { includeDocumentParagraphs: true },
  phrases: ['统一全文标题样式', '整理全文标题', '统一标书标题格式'],
  prepare(selection) {
    if (selection.kind !== 'text' || selection.targetId !== 'document-body')
      throw new HelperError('UNAVAILABLE', '请在 Word 文字文档中检查全文标题样式。');
    const paragraphs = selection.documentParagraphs;
    if (!paragraphs)
      throw new HelperError('UNAVAILABLE', '当前 Word 版本没有提供全文段落信息。');

    const analysis = analyzeDocumentStyles(paragraphs);
    const counts = { Heading1: 0, Heading2: 0, Heading3: 0 };
    for (const assignment of analysis.assignments) counts[assignment.style] += 1;
    const summary = [
      `全文共读取 ${paragraphs.length} 个段落。`,
      `可明确识别并规范的标题：一级 ${counts.Heading1} 个、二级 ${counts.Heading2} 个、三级 ${counts.Heading3} 个。`,
      `已有标题样式 ${analysis.existingHeadingCount} 个；列表和表格中的 ${analysis.skippedCount} 个段落会跳过。`,
      '预览按 Word 内置标题样式设置层级；正文、已有自定义样式、内容和编号均保留。应用后请检查分页和版面。',
      ...(['Heading1', 'Heading2', 'Heading3'] as const).flatMap((style) => {
        const examples = analysis.assignments
          .filter((assignment) => assignment.style === style)
          .slice(0, 3)
          .map((assignment) => `  · ${assignment.excerpt}`);
        return examples.length ? [`${headingNames[style]}示例：`, ...examples] : [];
      }),
    ].join('\n');

    if (!analysis.assignments.length) {
      return {
        summary,
        plan: {
          type: 'report',
          message:
            '没有发现可安全自动识别的编号标题。当前只按明确的编号样式判断；未编号的标题、自定义样式、列表和表格不会自动猜测。',
        },
      };
    }

    const plan: TextActionPlan = {
      type: 'text.documentStyles',
      assignments: analysis.assignments.map(({ paragraphIndex, style }) => ({
        paragraphIndex,
        style,
      })),
    };
    return { summary, plan, confirmLabel: '确认统一标题样式' };
  },
};
