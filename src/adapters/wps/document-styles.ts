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

type NativeObject = Record<string, unknown>;
type StyleEnums = Record<(typeof styleNames)[number][0], number>;

function object(value: unknown): value is NativeObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function unavailable(message: string): never {
  throw new HelperError('UNAVAILABLE', message);
}

/** Both layouts occur in synchronous bridges. Conflicting values are never guessed. */
function runtimeEnum(app: WpsApplication, group: string, key: string): number {
  const enums: unknown = app.Enum;
  if (!object(enums)) unavailable('当前 WPS 没有提供可验证的运行时枚举。');
  const flat = enums[key];
  const nested = enums[group];
  if (nested !== undefined && !object(nested)) unavailable('当前 WPS 的运行时枚举接口不完整。');
  const grouped = object(nested) ? nested[key] : undefined;
  if (
    (flat !== undefined && !Number.isSafeInteger(flat)) ||
    (grouped !== undefined && !Number.isSafeInteger(grouped)) ||
    (flat !== undefined && grouped !== undefined && flat !== grouped)
  )
    unavailable('当前 WPS 的运行时枚举无法可靠区分。');
  const value = grouped ?? flat;
  if (!Number.isSafeInteger(value)) unavailable('当前 WPS 没有提供可验证的运行时枚举。');
  return value as number;
}

function styleEnums(app: WpsApplication): StyleEnums {
  const values = styleNames.map(([, key]) => runtimeEnum(app, 'WdBuiltinStyle', key));
  if (new Set(values).size !== values.length)
    unavailable('当前 WPS 无法可靠区分正文与内置标题样式。');
  return Object.fromEntries(
    styleNames.map(([name], index) => [name, values[index]!]),
  ) as StyleEnums;
}

function readStyle(raw: string | number, styles: StyleEnums): BuiltInParagraphStyle {
  for (const [name] of styleNames) if (typeof raw === 'number' && raw === styles[name]) return name;
  if (typeof raw === 'number') return 'Other';
  // Keep internal punctuation/spacing in custom names. E.g. "Nor_mal" is not Normal.
  const normalized = raw.trim().toLowerCase();
  if (
    normalized === 'normal' ||
    normalized === 'wdstylenormal' ||
    normalized === '正文' ||
    normalized === '常规'
  )
    return 'Normal';
  const enumHeading = /^wdstyleheading([123])$/u.exec(normalized);
  if (enumHeading) return `Heading${enumHeading[1]}` as 'Heading1' | 'Heading2' | 'Heading3';
  const heading = /^(?:heading|标题)\s*([123])$/u.exec(normalized);
  if (heading) return `Heading${heading[1]}` as 'Heading1' | 'Heading2' | 'Heading3';
  if (normalized === 'title' || normalized === 'wdstyletitle' || normalized === '标题')
    return 'Title';
  if (normalized === 'subtitle' || normalized === 'wdstylesubtitle' || normalized === '副标题')
    return 'Subtitle';
  return 'Other';
}

function documentPath(document: NativeObject): string | undefined {
  const fullName = document.FullName;
  // Names alone cannot distinguish unsaved documents or equally named files.
  if (
    typeof fullName === 'string' &&
    /^(?:[a-z]:[\\/]|\\\\[^\\]+\\[^\\]+|\/|[a-z][a-z\d+.-]*:\/\/)/iu.test(fullName)
  )
    return fullName;
  return undefined;
}

function assertDocument(value: unknown, expected: WpsDocument): void {
  if (!object(value) || typeof value.Name !== 'string' || !value.Name.trim())
    unavailable('当前 WPS 无法确认全文范围所属的文档。');
  if (value === (expected as unknown as NativeObject)) return;
  const expectedPath = documentPath(expected as unknown as NativeObject);
  if (expectedPath && documentPath(value) === expectedPath) return;
  throw new HelperError('STALE_SELECTION', '全文段落不属于当前文档，请重新打开助手。');
}

function bounds(value: NativeObject): { start: number; end: number } {
  const start = value.Start;
  const end = value.End;
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    (start as number) < 0 ||
    (end as number) < (start as number)
  )
    unavailable('当前 WPS 无法确认全文段落的位置。');
  return { start: start as number, end: end as number };
}

function text(value: NativeObject): string {
  const result = value.Text;
  if (typeof result !== 'string') unavailable('当前 WPS 没有提供同步的全文文本接口。');
  return result;
}

function readRange(value: unknown, document: WpsDocument, mainStory: number) {
  if (!object(value)) unavailable('当前 WPS 无法读取完整段落范围。');
  assertDocument(value.Document, document);
  const position = bounds(value);
  if (value.StoryType !== mainStory) unavailable('当前 WPS 无法确认段落属于正文范围。');
  const nativeStyle = value.Style;
  if (!(
    (typeof nativeStyle === 'string' && !!nativeStyle.trim()) ||
    (typeof nativeStyle === 'number' && Number.isSafeInteger(nativeStyle))
  ))
    unavailable('当前 WPS 无法读取段落的原有样式。');
  const tables = value.Tables;
  const list = value.ListFormat;
  const tableCount = object(tables) ? tables.Count : undefined;
  const listType = object(list) ? list.ListType : undefined;
  if (
    !Number.isSafeInteger(tableCount) ||
    (tableCount as number) < 0 ||
    !Number.isSafeInteger(listType) ||
    (listType as number) < 0
  )
    unavailable('当前 WPS 无法确认全文中的表格和列表范围。');
  return {
    range: value as unknown as WpsRange,
    ...position,
    text: text(value),
    nativeStyle: nativeStyle as string | number,
    tableCount: tableCount as number,
    listType: listType as number,
  };
}

function readDocumentStyles(app: WpsApplication, document: WpsDocument): WpsDocumentStyleState {
  if (!object(document)) unavailable('当前 WPS 无法读取当前文档。');
  assertDocument(app.ActiveDocument, document);
  const identity = [document.Name, document.FullName];
  const styles = styleEnums(app);
  const none = runtimeEnum(app, 'WdListType', 'wdListNoNumbering');
  const mainStory = runtimeEnum(app, 'WdStoryType', 'wdMainTextStory');
  const content: unknown = document.Content;
  const collection = object(content) ? content.Paragraphs : undefined;
  const count = object(collection) ? collection.Count : undefined;
  if (
    !object(content) ||
    !object(collection) ||
    !Number.isSafeInteger(count) ||
    (count as number) < 1
  )
    unavailable('当前 WPS 无法读取全文段落。');
  if ((count as number) > MAX_DOCUMENT_PARAGRAPHS)
    throw new HelperError('SELECTION_TOO_LARGE', '全文段落超出本次安全检查上限。');
  const itemAt = collection.Item;
  if (typeof itemAt !== 'function') unavailable('当前 WPS 没有提供同步的段落集合接口。');
  assertDocument(content.Document, document);
  const position = bounds(content);
  if (content.StoryType !== mainStory) unavailable('当前 WPS 无法确认全文正文范围。');
  if (position.end - position.start > MAX_DOCUMENT_TEXT)
    throw new HelperError('SELECTION_TOO_LARGE', '全文字符数超出本次安全检查上限。');
  const contentText = text(content);
  if (contentText.length > MAX_DOCUMENT_TEXT)
    throw new HelperError('SELECTION_TOO_LARGE', '全文字符数超出本次安全检查上限。');

  const paragraphs: DocumentParagraphSnapshot[] = [];
  const targets: WpsStyleTarget[] = [];
  const duplicates = new Set<WpsRange>();
  let length = 0;
  let nextStart = position.start;

  for (let index = 1; index <= (count as number); index += 1) {
    const item: unknown = itemAt.call(collection, index);
    const source = readRange(object(item) ? item.Range : undefined, document, mainStory);
    const range = readRange(source.range.Duplicate, document, mainStory);
    if (range.range === source.range || duplicates.has(range.range))
      unavailable('当前 WPS 没有提供独立的段落副本范围。');
    duplicates.add(range.range);
    if (
      source.start !== range.start ||
      source.end !== range.end ||
      source.text !== range.text ||
      source.nativeStyle !== range.nativeStyle ||
      source.tableCount !== range.tableCount ||
      source.listType !== range.listType ||
      range.start !== nextStart ||
      range.end > position.end ||
      (range.start === range.end && !(count === 1 && contentText === ''))
    )
      throw new HelperError('STALE_SELECTION', '全文段落范围已改变，请重新预览。');
    length += range.text.length;
    if (length > MAX_DOCUMENT_TEXT)
      throw new HelperError('SELECTION_TOO_LARGE', '全文字符数超出本次安全检查上限。');
    nextStart = range.end;
    paragraphs.push({
      index: index - 1,
      text: range.text,
      style: readStyle(range.nativeStyle, styles),
      isListItem: range.listType !== none,
      // The end-of-cell marker remains a table signal even on incomplete bridges.
      inTable: range.tableCount > 0 || range.text.includes('\u0007'),
    });
    targets.push({ range: range.range, nativeStyle: range.nativeStyle });
  }

  assertDocument(app.ActiveDocument, document);
  if (
    collection.Count !== count ||
    nextStart !== position.end ||
    paragraphs.map((paragraph) => paragraph.text).join('') !== contentText ||
    document.Name !== identity[0] ||
    document.FullName !== identity[1] ||
    content.Start !== position.start ||
    content.End !== position.end ||
    content.StoryType !== mainStory ||
    text(content) !== contentText
  )
    throw new HelperError('STALE_SELECTION', '全文内容在读取时发生变化，请重新预览。');
  return { paragraphs, targets };
}

/**
 * Read-only probe of the current synchronous API and document, within the safety limits.
 * A positive result does not verify native Style writes or real-machine compatibility.
 * Every later read must validate the current document again.
 */
export function supportsWpsDocumentStyles(app: WpsApplication): boolean {
  try {
    readWpsDocumentStyles(app, app.ActiveDocument!);
    return true;
  } catch {
    return false;
  }
}

export function readWpsDocumentStyles(
  app: WpsApplication,
  document: WpsDocument,
): WpsDocumentStyleState {
  try {
    return readDocumentStyles(app, document);
  } catch (error) {
    if (error instanceof HelperError) throw error;
    throw new HelperError('UNAVAILABLE', '当前 WPS 的全文样式接口无法安全读取。');
  }
}

export function nativeWpsStyle(
  app: WpsApplication,
  style: 'Heading1' | 'Heading2' | 'Heading3',
): number {
  try {
    if (style !== 'Heading1' && style !== 'Heading2' && style !== 'Heading3')
      throw new HelperError('INVALID_PLAN', '只能使用明确的内置一级、二级或三级标题样式。');
    return styleEnums(app)[style];
  } catch (error) {
    if (error instanceof HelperError) throw error;
    throw new HelperError('UNAVAILABLE', '当前 WPS 无法读取内置标题样式枚举。');
  }
}

/**
 * Apply-side guard for a paragraph that was Normal when the preview was made.
 * WPS may return either the runtime enum value or a localized style name here.
 */
export function isWpsNormalStyle(app: WpsApplication, raw: string | number): boolean {
  return readStyle(raw, styleEnums(app)) === 'Normal';
}
