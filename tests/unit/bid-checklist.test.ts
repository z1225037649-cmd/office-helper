import { describe, expect, it } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import { FeatureRegistry } from '../../src/core/registry';
import { TaskRunner } from '../../src/core/runner';
import type { DocumentParagraphSnapshot, SelectionSnapshot } from '../../src/core/contracts';
import { features } from '../../src/features';
import { bidChecklist } from '../../src/features/bid-checklist/feature';

describe('标书要求候选清单工作流', () => {
  it('注册为文字文档只读报告，保留来源段落、原文和偏移', async () => {
    const host = new MockAdapter('text');
    const documentParagraphs: DocumentParagraphSnapshot[] = [
      {
        index: 0,
        text: '投标人必须提供有效营业执照。',
        style: 'Normal',
        isListItem: false,
        inTable: false,
      },
      {
        index: 1,
        text: '响应方案应满足技术参数。',
        style: 'Normal',
        isListItem: false,
        inTable: false,
      },
    ];
    const selection: SelectionSnapshot = {
      documentId: 'demo-notice',
      documentName: '招标文件.docx',
      kind: 'text',
      targetId: 'document-body',
      targetLabel: '全文（2 个段落）',
      revision: JSON.stringify(documentParagraphs),
      documentParagraphs,
    };
    const readSelection = host.readSelection.bind(host);
    host.readSelection = async (options) =>
      options?.includeDocumentParagraphs ? structuredClone(selection) : readSelection(options);
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    const before = await host.readSelection({ includeDocumentParagraphs: true });
    const prepared = await runner.prepare('document.bidChecklist');

    expect(prepared.changesDocument).toBe(false);
    expect(prepared.summary).toContain('人工确认');
    const result = await runner.execute(prepared.id);

    expect(result.canUndo).toBe(false);
    expect(result.message).toContain('不是完整性或合规结论');
    expect(result.findings).toEqual([
      expect.objectContaining({
        label: '资格与资质 · 待核对',
        sourceExcerpt: '投标人必须提供有效营业执照。',
        category: 'qualification',
        row: 0,
        column: 0,
        startOffset: 0,
        endOffset: '投标人必须提供有效营业执照。'.length,
      }),
      expect.objectContaining({
        label: '技术要求 · 待核对',
        sourceExcerpt: '响应方案应满足技术参数。',
        category: 'technical',
        row: 1,
      }),
    ]);
    expect(await host.readSelection({ includeDocumentParagraphs: true })).toEqual(before);
  });

  it('目录和本地短句可以找到入口', async () => {
    expect(features).toContain(bidChecklist);
    const { LocalIntentProvider } = await import('../../src/services/intent');
    await expect(new LocalIntentProvider().resolve('请检查标书要求', features)).resolves.toEqual({
      type: 'match',
      featureId: 'document.bidChecklist',
    });
  });

  it('对不带全文快照的文字范围不猜测或扫描选区以外内容', () => {
    expect(() =>
      bidChecklist.prepare({
        documentId: 'doc',
        documentName: '稿件.docx',
        kind: 'text',
        targetId: 'selection',
        targetLabel: '选中文字',
        revision: '1',
        text: '必须提供证明。',
      }),
    ).toThrow('请在 Word');
  });
});
