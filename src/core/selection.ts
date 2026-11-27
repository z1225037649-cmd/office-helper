import type { SelectionSnapshot } from './contracts';
import { HelperError } from './errors';

export const MAX_SELECTION_CELLS = 5000;
export const MAX_SELECTION_TEXT = 20000;
export const MAX_SELECTION_PARAGRAPHS = 20;
export const MAX_DOCUMENT_PARAGRAPHS = 1500;
export const MAX_DOCUMENT_TEXT = 120000;

export function selectionKey(selection: SelectionSnapshot): string {
  return JSON.stringify([
    selection.documentId,
    selection.kind,
    selection.targetId,
    selection.revision,
    selection.text,
    selection.fontSize,
    selection.bold,
    selection.paragraphs,
    selection.documentParagraphs,
    selection.values,
    selection.cellTypes,
    selection.rowScope,
    selection.sheetLayout,
    selection.printSettings,
  ]);
}

export function assertSameSelection(expected: SelectionSnapshot, actual: SelectionSnapshot): void {
  if (selectionKey(expected) !== selectionKey(actual)) {
    throw new HelperError('STALE_SELECTION', '文件内容或选中范围变了，请重新预览后再操作。');
  }
}

export function assertSelectionLimit(selection: SelectionSnapshot): void {
  const cellCount = selection.values?.reduce((count, row) => count + row.length, 0) ?? 0;
  const paragraphTextLength = selection.paragraphs?.reduce((n, p) => n + p.text.length, 0) ?? 0;
  const documentTextLength =
    selection.documentParagraphs?.reduce((n, paragraph) => n + paragraph.text.length, 0) ?? 0;
  if (
    cellCount > MAX_SELECTION_CELLS ||
    (selection.text?.length ?? 0) > MAX_SELECTION_TEXT ||
    (selection.paragraphs?.length ?? 0) > MAX_SELECTION_PARAGRAPHS ||
    paragraphTextLength > MAX_SELECTION_TEXT ||
    (selection.documentParagraphs?.length ?? 0) > MAX_DOCUMENT_PARAGRAPHS ||
    documentTextLength > MAX_DOCUMENT_TEXT
  ) {
    throw new HelperError(
      'SELECTION_TOO_LARGE',
      selection.documentParagraphs
        ? '全文内容较多，本次最多检查 1,500 个段落或 120,000 个字符。'
        : '选中的内容有点多，请缩小范围后再试。',
    );
  }
}
