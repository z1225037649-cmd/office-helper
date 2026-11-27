import { describe, expect, it } from 'vitest';
import type { DocumentParagraphSnapshot, SelectionSnapshot } from '../../core/contracts';
import { bidChecklist } from './feature';

function prepareParagraph(text: string) {
  const paragraph: DocumentParagraphSnapshot = {
    index: 0,
    text,
    style: 'Normal',
    isListItem: false,
    inTable: false,
  };
  const selection: SelectionSnapshot = {
    documentId: 'long-notice',
    documentName: '招标文件.docx',
    kind: 'text',
    targetId: 'document-body',
    targetLabel: '全文',
    revision: 'test-revision',
    documentParagraphs: [paragraph],
  };
  const prepared = bidChecklist.prepare(selection);
  if (prepared.plan.type !== 'report') throw new Error('Expected a read-only report');
  return { prepared, findings: prepared.plan.findings ?? [] };
}

describe('bidChecklist finding display boundaries', () => {
  it('caps excerpts without splitting a surrogate pair and preserves full source offsets', () => {
    const text = `${'x'.repeat(499)}😀 必须提交投标文件。`;
    const { findings } = prepareParagraph(text);
    const [finding] = findings;

    expect(findings).toHaveLength(1);
    expect(finding?.sourceExcerpt).toBe('x'.repeat(499));
    expect(finding?.sourceExcerpt).not.toContain('\ud83d');
    expect(finding?.startOffset).toBe(0);
    expect(finding?.endOffset).toBe(text.length);
    expect(finding?.description).toContain('显示片段已截短');
  });

  it('limits displayed candidates while reporting the full candidate count', () => {
    const text = Array.from({ length: 205 }, (_, index) => `${index + 1}. 必须提供证明。`).join(
      '\n',
    );
    const { prepared, findings } = prepareParagraph(text);

    expect(findings).toHaveLength(200);
    expect(prepared.summary).toContain('共发现 205 条候选要求');
    expect(prepared.summary).toContain('只显示前 200 条');
    expect(findings[0]?.startOffset).toBe(0);
    expect(findings[199]?.endOffset).toBe(text.split('\n').slice(0, 200).join('\n').length);
  });
});
