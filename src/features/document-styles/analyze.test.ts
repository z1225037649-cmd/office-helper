import { describe, expect, it } from 'vitest';
import type { DocumentParagraphSnapshot } from '../../core/contracts';
import { analyzeDocumentStyles } from './analyze';

function paragraph(
  index: number,
  text: string,
  overrides: Partial<DocumentParagraphSnapshot> = {},
): DocumentParagraphSnapshot {
  return {
    index,
    text,
    style: 'Normal',
    isListItem: false,
    inTable: false,
    ...overrides,
  };
}

describe('analyzeDocumentStyles', () => {
  it('maps strict Chinese and decimal heading patterns to built-in styles', () => {
    expect(
      analyzeDocumentStyles([
        paragraph(0, '第一章 项目概况'),
        paragraph(1, '一、项目说明'),
        paragraph(2, '（一）服务范围'),
        paragraph(3, '2.1.1 技术要求'),
        paragraph(4, '普通正文内容。'),
      ]).assignments,
    ).toEqual([
      { paragraphIndex: 0, style: 'Heading1', excerpt: '第一章 项目概况' },
      { paragraphIndex: 1, style: 'Heading1', excerpt: '一、项目说明' },
      { paragraphIndex: 2, style: 'Heading2', excerpt: '（一）服务范围' },
      { paragraphIndex: 3, style: 'Heading3', excerpt: '2.1.1 技术要求' },
    ]);
  });

  it('preserves existing headings, custom styles, lists, tables, and unnumbered prose', () => {
    const analysis = analyzeDocumentStyles([
      paragraph(0, '第一章 已标记标题', { style: 'Heading2' }),
      paragraph(1, '第二章 自定义格式', { style: 'Other' }),
      paragraph(2, '第三章 列表项', { isListItem: true }),
      paragraph(3, '第四章 表格标题', { inTable: true }),
      paragraph(4, '未编号标题'),
      paragraph(5, '正文。'),
    ]);

    expect(analysis.assignments).toEqual([]);
    expect(analysis.existingHeadingCount).toBe(1);
    expect(analysis.skippedCount).toBe(2);
  });
});
