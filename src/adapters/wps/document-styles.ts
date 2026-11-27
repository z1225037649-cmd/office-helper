import type { BuiltInParagraphStyle, DocumentParagraphSnapshot } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { MAX_DOCUMENT_PARAGRAPHS, MAX_DOCUMENT_TEXT } from '../../core/selection';
import type { WpsApplication, WpsDocument, WpsRange } from './bridge';

interface WpsStyleTarget {
  range: WpsRange;
  nativeStyle: string | number;
}

export interface WpsDocumentStyleState {
  paragraphs: DocumentParagraphSnapshot[];
  targets: WpsStyleTarget[];
}

const styleNames = [
  ['Normal', 'wdStyleNormal'],
  ['Heading1', 'wdStyleHeading1'],
  ['Heading2', 'wdStyleHeading2'],
  ['Heading3', 'wdStyleHeading3'],
] as const;

function styleEnum(app: WpsApplication, style: 'Heading1' | 'Heading2' | 'Heading3'): number {
  const key = styleNames.find(([name]) => name === style)?.[1];
  const value = key ? app.Enum?.[key] : undefined;
  if (!Number.isInteger(value))
    throw new HelperError('UNAVAILABLE', '当前 WPS 没有提供可验证的内置标题样式编号。');
  return value!;
}

function readStyle(raw: string | number, app: WpsApplication): BuiltInParagraphStyle {
  for (const [name, key] of styleNames)
    if (typeof raw === 'number' && raw === app.Enum?.[key]) return name;
  const normalized = String(raw).replace(/[\s_-]/gu, '').toLocaleLowerCase();
  if (normalized === 'normal' || normalized === '正文' || normalized === '常规') return 'Normal';
  if (normalized === 'heading1' || normalized === '标题1') return 'Heading1';
  if (normalized === 'heading2' || normalized === '标题2') return 'Heading2';
  if (normalized === 'heading3' || normalized === '标题3') return 'Heading3';
  if (normalized === 'title' || normalized === '标题') return 'Title';
  if (normalized === 'subtitle' || normalized === '副标题') return 'Subtitle';
  return 'Other';
}

function noListStyle(app: WpsApplication): number {
  const value = app.Enum?.WdListType?.wdListNoNumbering ?? app.Enum?.wdListNoNumbering;
  if (!Number.isInteger(value))
    throw new HelperError('UNAVAILABLE', '当前 WPS 无法可靠区分列表段落。');
  return value!;
}

export function supportsWpsDocumentStyles(app: WpsApplication): boolean {
  try {
    const constants = [
      app.Enum?.wdStyleNormal,
      app.Enum?.wdStyleHeading1,
      app.Enum?.wdStyleHeading2,
      app.Enum?.wdStyleHeading3,
    ];
    const noList = app.Enum?.WdListType?.wdListNoNumbering ?? app.Enum?.wdListNoNumbering;
    const content = app.ActiveDocument?.Content;
    const first = content?.Paragraphs?.Count ? content.Paragraphs.Item(1).Range : undefined;
    return (
      constants.every(Number.isInteger) &&
      new Set(constants).size === constants.length &&
      Number.isInteger(noList) &&
      !!first &&
      typeof first.Style !== 'undefined' &&
      Number.isInteger(first.Tables?.Count) &&
      Number.isInteger(first.ListFormat?.ListType)
    );
  } catch {
    return false;
  }
}

export function readWpsDocumentStyles(
  app: WpsApplication,
  document: WpsDocument,
): WpsDocumentStyleState {
  const content = document.Content;
  const collection = content?.Paragraphs;
  if (!content || !collection || !Number.isInteger(collection.Count) || collection.Count < 1)
    throw new HelperError('UNAVAILABLE', '当前 WPS 无法读取全文段落。');
  if (collection.Count > MAX_DOCUMENT_PARAGRAPHS)
    throw new HelperError('SELECTION_TOO_LARGE', '全文段落超出本次安全检查上限。');
  const none = noListStyle(app);
  const paragraphs: DocumentParagraphSnapshot[] = [];
  const targets: WpsStyleTarget[] = [];
  let length = 0;

  for (let index = 1; index <= collection.Count; index += 1) {
    const range = collection.Item(index).Range.Duplicate;
    const nativeStyle = range.Style;
    const tableCount = range.Tables?.Count;
    const listType = range.ListFormat?.ListType;
    if (
      !(
        (typeof nativeStyle === 'string' && !!nativeStyle.trim()) ||
        (typeof nativeStyle === 'number' && Number.isFinite(nativeStyle))
      )
    )
      throw new HelperError('UNAVAILABLE', '当前 WPS 无法读取段落的原有样式。');
    if (!Number.isInteger(tableCount) || !Number.isInteger(listType))
      throw new HelperError('UNAVAILABLE', '当前 WPS 无法确认全文中的表格和列表范围。');
    if (range.Document && (range.Document.FullName || range.Document.Name) !== (document.FullName || document.Name))
      throw new HelperError('STALE_SELECTION', '全文段落不属于当前文档，请重新打开助手。');
    const text = String(range.Text ?? '');
    length += text.length;
    paragraphs.push({
      index: index - 1,
      text,
      style: readStyle(nativeStyle, app),
      isListItem: listType! !== none,
      inTable: tableCount! > 0,
    });
    targets.push({ range, nativeStyle });
  }

  if (length > MAX_DOCUMENT_TEXT)
    throw new HelperError('SELECTION_TOO_LARGE', '全文字符数超出本次安全检查上限。');
  return { paragraphs, targets };
}

export function nativeWpsStyle(
  app: WpsApplication,
  style: 'Heading1' | 'Heading2' | 'Heading3',
): number {
  return styleEnum(app, style);
}
