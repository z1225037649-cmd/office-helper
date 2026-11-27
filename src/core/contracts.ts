export type HostId = 'mock' | 'office' | 'wps';
export type DocumentKind = 'text' | 'sheet';
export type Capability =
  | 'selection.read'
  | 'text.resize'
  | 'text.bold'
  | 'text.documentStyles'
  | 'paragraph.center'
  | 'paragraph.spacing'
  | 'paragraph.indent'
  | 'sheet.autofit'
  | 'sheet.wrap'
  | 'sheet.header'
  | 'sheet.print'
  | 'document.saveCopy'
  | 'document.exportPdf'
  | 'document.notice'
  | 'selection.locate'
  | 'history.undo';
export type CellValue = string | number | boolean | null;
export type CellType = 'empty' | 'number' | 'text' | 'boolean' | 'error' | 'unknown';
export type ExportScope = 'document' | 'selection' | 'activeSheet';

/** Extra native reads are opt-in so ordinary selection and sum keep their existing path. */
export interface SelectionReadOptions {
  includeCellTypes?: boolean;
  includeSheetLayout?: boolean;
  includePrintSettings?: boolean;
  includeDocumentParagraphs?: boolean;
}

/** Native units are retained verbatim for conflict-safe restoration. */
export interface SheetLayoutSnapshot {
  columnWidths: number[];
  widthUnit: 'points' | 'characters';
  wrapText: boolean[][];
  headerBold: boolean[];
  merged: boolean;
  protected: boolean;
  firstRow: number;
  firstColumn: number;
  sheetName: string;
}

export interface PrintSettingsSnapshot {
  available: boolean;
  reason?: string;
  printArea: string;
  orientation: string;
  zoom: number | false;
  fitWide: number;
  fitTall: number;
  /** A measured content width, not an invented page-count estimate. */
  contentWidthPoints?: number;
}

export interface NoticeFields {
  title: string;
  time: string;
  location: string;
  body: string;
  contact: string;
}

export interface DuplicateOptions {
  keyColumns: number[];
  requiredColumns: number[];
  trimWhitespace: boolean;
  ignoreCase: boolean;
  hasHeader: boolean;
}

/** Row/column are zero-based offsets inside the original inspected range. */
export interface ReportFinding {
  id: string;
  label: string;
  description: string;
  row: number;
  column: number;
  /** Text-document finding metadata; offsets are UTF-16 code units within one paragraph. */
  sourceExcerpt?: string;
  category?: string;
  status?: string;
  startOffset?: number;
  endOffset?: number;
}

export interface ParagraphSnapshot {
  text: string;
  alignment: string;
  /** Office.js uses 12 for single spacing, 18 for 1.5, etc. */
  lineSpacing: number;
  /** Only explicit multiple spacing is currently safe to restore through this API. */
  spacingMode: 'multiple' | 'unsupported';
  firstLineIndent: number;
  /** Native hosts may explicitly reject an indent whose original units cannot be restored. */
  firstLineIndentRestorable?: boolean;
  leftIndent: number;
  rightIndent: number;
  isListItem: boolean;
}

export type BuiltInParagraphStyle =
  'Normal' | 'Title' | 'Subtitle' | 'Heading1' | 'Heading2' | 'Heading3' | 'Other';

/** Lightweight, serializable paragraph data used by explicit whole-document tasks. */
export interface DocumentParagraphSnapshot {
  index: number;
  text: string;
  style: BuiltInParagraphStyle;
  isListItem: boolean;
  inTable: boolean;
}

export interface HostInfo {
  id: HostId;
  label: string;
  kind: DocumentKind;
  capabilities: readonly Capability[];
  notice?: string;
}

/** 只传普通数据。Office/WPS 的代理对象不得越过适配层。 */
export interface SelectionSnapshot {
  documentId: string;
  documentName: string;
  kind: DocumentKind;
  targetId: string;
  targetLabel: string;
  revision: string;
  text?: string;
  fontSize?: number;
  bold?: boolean;
  paragraphs?: ParagraphSnapshot[];
  documentParagraphs?: DocumentParagraphSnapshot[];
  values?: CellValue[][];
  /** Native cell types distinguish spreadsheet errors from text or numeric error codes. */
  cellTypes?: CellType[][];
  /** All cells in a contiguous range, including hidden and filtered rows. */
  rowScope?: 'all';
  sheetLayout?: SheetLayoutSnapshot;
  printSettings?: PrintSettingsSnapshot;
}

export type TextActionPlan =
  | { type: 'text.resize'; size: number }
  | { type: 'text.bold'; bold: boolean }
  | {
      type: 'text.documentStyles';
      assignments: { paragraphIndex: number; style: 'Heading1' | 'Heading2' | 'Heading3' }[];
    }
  | { type: 'paragraph.center' }
  | { type: 'paragraph.spacing'; multiple: number }
  | { type: 'paragraph.indent'; points: number };

export type ExtendedActionPlan =
  | { type: 'sheet.autofit' }
  | { type: 'sheet.wrap' }
  | { type: 'sheet.header' }
  | { type: 'sheet.print' }
  | { type: 'document.saveCopy'; fileName: string; folder?: string }
  | { type: 'document.exportPdf'; fileName: string; folder?: string; scope?: ExportScope }
  | { type: 'document.notice'; fields: NoticeFields };

export type ActionPlan =
  | { type: 'report'; message: string; findings?: ReportFinding[]; outcome?: 'guidance' }
  | TextActionPlan
  | ExtendedActionPlan;

export function isExtendedPlan(plan: ActionPlan): plan is ExtendedActionPlan {
  return plan.type.startsWith('sheet.') || plan.type.startsWith('document.');
}

/** Explicit, safely parsed parameters supplied by the user for a feature. */
export interface FeatureParameters {
  fontSize?: number;
  bold?: boolean;
  lineSpacingMultiple?: number;
  firstLineIndentPoints?: number;
  duplicates?: DuplicateOptions;
  notice?: NoticeFields;
  fileName?: string;
  folder?: string;
  exportScope?: ExportScope;
}

export interface ChangeReceipt {
  token: string;
  message: string;
  canUndo: boolean;
  outcome?: 'completed' | 'guidance' | 'cancelled' | 'download';
  location?: string;
  actions?: ResultAction[];
}

/** Opaque host-owned identifiers; the UI never opens caller-supplied paths. */
export interface ResultAction {
  id: string;
  label: string;
}

export interface HostAdapter {
  readonly info: HostInfo;
  readSelection(options?: SelectionReadOptions): Promise<SelectionSnapshot>;
  /** 原生适配器必须在写入前再次核对目标和快照，拒绝过期任务。 */
  apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt>;
  undo(receipt: ChangeReceipt): Promise<void>;
  locate?(selection: SelectionSnapshot, row: number, column: number): Promise<void>;
  openResult?(actionId: string): Promise<void>;
  dispose(): void;
}

export interface FeatureDefinition {
  id: string;
  title: string;
  description: string;
  icon: 'text' | 'sum' | 'inspect';
  kinds: readonly DocumentKind[];
  requires: readonly Capability[];
  phrases: readonly string[];
  readOptions?: SelectionReadOptions;
  /** A parameterized button focuses the request field instead of inventing a value. */
  inputPrompt?: { prefix: string; message: string };
  form?: 'duplicates' | 'notice' | 'save';
  prepare(
    selection: SelectionSnapshot,
    parameters?: FeatureParameters,
  ): { summary: string; plan: ActionPlan; confirmLabel?: string };
}

export interface PreparedTask {
  id: string;
  title: string;
  target: string;
  summary: string;
  changesDocument: boolean;
  confirmLabel?: string;
}

export interface TaskResult {
  message: string;
  canUndo: boolean;
  outcome?: ChangeReceipt['outcome'];
  location?: string;
  findings?: ReportFinding[];
  actions?: ResultAction[];
}

export interface ActivityEntry {
  id: string;
  featureId: string;
  title: string;
  status: 'completed' | 'restored';
  time: number;
}
