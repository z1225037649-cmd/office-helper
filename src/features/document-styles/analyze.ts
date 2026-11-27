import type { DocumentParagraphSnapshot } from '../../core/contracts';

export type HeadingStyle = 'Heading1' | 'Heading2' | 'Heading3';

export interface HeadingStyleAssignment {
  paragraphIndex: number;
  style: HeadingStyle;
  excerpt: string;
}

export interface DocumentStyleAnalysis {
  assignments: HeadingStyleAssignment[];
  skippedCount: number;
  existingHeadingCount: number;
}

const headingOne = /^(?:第[〇零一二三四五六七八九十百千万\d]{1,12}[章节部分篇]|[一二三四五六七八九十]+[、.．])\s*\S/u;
const headingTwo = /^(?:（[一二三四五六七八九十]+）|\([一二三四五六七八九十]+\)|\d+\.\d+(?!\.)\s+)\s*\S/u;
const headingThree = /^\d+\.\d+\.\d+(?:\.\d+)?\s+\S/u;

function isHeadingStyle(style: DocumentParagraphSnapshot['style']): boolean {
  return style === 'Heading1' || style === 'Heading2' || style === 'Heading3';
}

function inferHeading(text: string): HeadingStyle | undefined {
  const content = text.trim();
  if (headingThree.test(content)) return 'Heading3';
  if (headingTwo.test(content)) return 'Heading2';
  if (headingOne.test(content)) return 'Heading1';
  return undefined;
}

/**
 * Recommends only strict, numbered heading candidates that currently use Normal style.
 * Table, list, empty, custom-style, and already styled paragraphs are preserved.
 */
export function analyzeDocumentStyles(
  paragraphs: readonly DocumentParagraphSnapshot[],
): DocumentStyleAnalysis {
  const assignments: HeadingStyleAssignment[] = [];
  let skippedCount = 0;
  let existingHeadingCount = 0;

  for (const paragraph of paragraphs) {
    if (isHeadingStyle(paragraph.style)) existingHeadingCount += 1;
    if (paragraph.inTable || paragraph.isListItem) {
      skippedCount += 1;
      continue;
    }
    if (!paragraph.text.trim() || paragraph.style !== 'Normal') continue;
    const style = inferHeading(paragraph.text);
    if (style) {
      assignments.push({
        paragraphIndex: paragraph.index,
        style,
        excerpt: paragraph.text.trim().slice(0, 100),
      });
    }
  }

  return { assignments, skippedCount, existingHeadingCount };
}
