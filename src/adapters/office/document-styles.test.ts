import { afterEach, describe, expect, it, vi } from 'vitest';
import { OfficeAdapter } from './adapter';
import type { BuiltInParagraphStyle } from '../../core/contracts';

function setup() {
  const source = [
    {
      text: '第一章 项目概况',
      style: 'Normal' as BuiltInParagraphStyle,
      isListItem: false,
      tableNestingLevel: 0,
    },
    {
      text: '一、项目说明',
      style: 'Normal' as BuiltInParagraphStyle,
      isListItem: false,
      tableNestingLevel: 0,
    },
    {
      text: '（一）服务范围',
      style: 'Normal' as BuiltInParagraphStyle,
      isListItem: false,
      tableNestingLevel: 0,
    },
    {
      text: '正文内容。',
      style: 'Normal' as BuiltInParagraphStyle,
      isListItem: false,
      tableNestingLevel: 0,
    },
    {
      text: '第四章 表格内标题',
      style: 'Normal' as BuiltInParagraphStyle,
      isListItem: false,
      tableNestingLevel: 1,
    },
    {
      text: '第五章 列表内标题',
      style: 'Normal' as BuiltInParagraphStyle,
      isListItem: true,
      tableNestingLevel: 0,
    },
  ];
  const paragraphs = source.map((item) => ({
    get text() {
      return item.text;
    },
    get styleBuiltIn() {
      return item.style;
    },
    set styleBuiltIn(style: BuiltInParagraphStyle) {
      item.style = style;
    },
    get isListItem() {
      return item.isListItem;
    },
    get tableNestingLevel() {
      return item.tableNestingLevel;
    },
    load: vi.fn(),
  }));
  const collection = { items: paragraphs, load: vi.fn() };
  const wordContext = {
    document: { body: { paragraphs: collection } },
    sync: vi.fn(async () => undefined),
  };
  const run = vi.fn(async (...args: unknown[]) => {
    const callback = args.at(-1) as (requestContext: typeof wordContext) => Promise<unknown>;
    return callback(wordContext);
  });
  vi.stubGlobal('Office', {
    context: {
      document: { url: 'https://example.test/bid.docx' },
    },
  });
  vi.stubGlobal('Word', { run });
  const host = new OfficeAdapter('text', true);
  return { host, source, run };
}

afterEach(() => vi.unstubAllGlobals());

describe('Office 全文标题样式', () => {
  it('仅按确认计划给安全候选套用内置标题，并能冲突安全恢复', async () => {
    const { host, source, run } = setup();
    const before = await host.readSelection({ includeDocumentParagraphs: true });
    const plan = {
      type: 'text.documentStyles' as const,
      assignments: [
        { paragraphIndex: 0, style: 'Heading1' as const },
        { paragraphIndex: 1, style: 'Heading1' as const },
        { paragraphIndex: 2, style: 'Heading2' as const },
      ],
    };

    const receipt = await host.apply(plan, before);

    expect(receipt.canUndo).toBe(true);
    expect(source.map(({ style }) => style)).toEqual([
      'Heading1',
      'Heading1',
      'Heading2',
      'Normal',
      'Normal',
      'Normal',
    ]);
    await host.undo(receipt);
    expect(source.every(({ style }) => style === 'Normal')).toBe(true);
    expect(run).toHaveBeenCalled();
  });

  it('确认前正文变化时拒绝写入', async () => {
    const { host, source } = setup();
    const before = await host.readSelection({ includeDocumentParagraphs: true });
    source[0]!.text = '第一章 新内容';

    await expect(
      host.apply(
        { type: 'text.documentStyles', assignments: [{ paragraphIndex: 0, style: 'Heading1' }] },
        before,
      ),
    ).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(source[0]!.style).toBe('Normal');
  });

  it('有后续全文修改时拒绝恢复，不覆盖新内容', async () => {
    const { host, source } = setup();
    const before = await host.readSelection({ includeDocumentParagraphs: true });
    const receipt = await host.apply(
      { type: 'text.documentStyles', assignments: [{ paragraphIndex: 0, style: 'Heading1' }] },
      before,
    );
    source[3]!.text = '用户后来编辑的正文。';

    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(source[0]!.style).toBe('Heading1');
    expect(source[3]!.text).toBe('用户后来编辑的正文。');
  });
});
