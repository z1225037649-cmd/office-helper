import type { FeatureDefinition, ReportFinding } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { analyzeBidChecklist } from './analyze';

const categoryLabels = {
  qualification: '资格与资质',
  technical: '技术要求',
  commercial: '商务与报价',
  submission: '文件提交',
  deadline: '时间节点',
  other: '其他要求',
} as const;

const displayLimit = 200;
const excerptLimit = 500;

function excerptForDisplay(text: string): string {
  let end = Math.min(text.length, excerptLimit);
  if (
    end < text.length &&
    end > 0 &&
    text.charCodeAt(end - 1) >= 0xd800 &&
    text.charCodeAt(end - 1) <= 0xdbff &&
    text.charCodeAt(end) >= 0xdc00 &&
    text.charCodeAt(end) <= 0xdfff
  ) {
    end -= 1;
  }
  return text.slice(0, end);
}

export const bidChecklist: FeatureDefinition = {
  id: 'document.bidChecklist',
  title: '检查标书要求',
  description: '扫描 Word 正文中疑似资格、技术、商务、提交和时间要求，保留原文供人工核对',
  icon: 'inspect',
  kinds: ['text'],
  requires: ['selection.read', 'text.documentStyles'],
  readOptions: { includeDocumentParagraphs: true },
  phrases: ['检查标书要求', '提取标书要求', '整理投标要求', '生成投标核对清单'],
  prepare(selection) {
    if (selection.kind !== 'text' || selection.targetId !== 'document-body')
      throw new HelperError('UNAVAILABLE', '请在 Word 文字文档中检查标书要求。');
    const paragraphs = selection.documentParagraphs;
    if (!paragraphs) throw new HelperError('UNAVAILABLE', '当前 Word 版本没有提供全文段落信息。');

    const findings: ReportFinding[] = [];
    let candidateCount = 0;
    for (const paragraph of paragraphs) {
      const candidates = analyzeBidChecklist(paragraph.text);
      candidateCount += candidates.length;
      for (const candidate of candidates) {
        if (findings.length >= displayLimit) continue;
        const isTruncated = candidate.originalExcerpt.length > excerptLimit;
        findings.push({
          id: `bid-requirement-${paragraph.index}-${candidate.startOffset}-${findings.length}`,
          label: `${categoryLabels[candidate.category]} · 待核对`,
          description: `规则识别出的候选条目，不代表最终要求结论。源文第 ${paragraph.index + 1} 段，段内字符偏移 ${candidate.startOffset}–${candidate.endOffset}。${isTruncated ? '显示片段已截短，请以原文为准。' : ''}`,
          row: paragraph.index,
          column: 0,
          sourceExcerpt: excerptForDisplay(candidate.originalExcerpt),
          category: candidate.category,
          status: candidate.status,
          startOffset: candidate.startOffset,
          endOffset: candidate.endOffset,
        });
      }
    }

    const displaySummary =
      candidateCount > displayLimit
        ? `共发现 ${candidateCount} 条候选要求，结果只显示前 ${displayLimit} 条。`
        : `共发现 ${candidateCount} 条候选要求，结果已全部显示。`;
    return {
      summary: `只读扫描「${selection.documentName}」正文 ${paragraphs.length} 个段落。按中英文常见要求词识别候选内容，原文不作改写；分类和“待核对”状态均需人工确认。${displaySummary}`,
      plan: {
        type: 'report',
        message: `${displaySummary}这是候选核对清单，不是完整性或合规结论；请对照采购文件逐条确认要求、响应内容和证明材料。正文之外的页眉、页脚、文本框和附件不会纳入本次扫描。`,
        findings,
      },
    };
  },
};
