export type HostId = 'mock' | 'office' | 'wps';
export type DocumentKind = 'text' | 'sheet';
export type Capability =
  | 'selection.read'
  | 'text.resize'
  | 'text.bold'
  | 'paragraph.center'
  | 'paragraph.spacing'
  | 'paragraph.indent'
  | 'history.undo';
export type CellValue = string | number | boolean | null;

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
  values?: CellValue[][];
}

export type TextActionPlan =
  | { type: 'text.resize'; size: number }
  | { type: 'text.bold'; bold: boolean }
  | { type: 'paragraph.center' }
  | { type: 'paragraph.spacing'; multiple: number }
  | { type: 'paragraph.indent'; points: number };

export type ActionPlan = { type: 'report'; message: string } | TextActionPlan;

/** Explicit, safely parsed parameters supplied by the user for a feature. */
export interface FeatureParameters {
  fontSize?: number;
  bold?: boolean;
  lineSpacingMultiple?: number;
  firstLineIndentPoints?: number;
}

export interface ChangeReceipt {
  token: string;
  message: string;
  canUndo: boolean;
}

export interface HostAdapter {
  readonly info: HostInfo;
  readSelection(): Promise<SelectionSnapshot>;
  /** 原生适配器必须在写入前再次核对目标和快照，拒绝过期任务。 */
  apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt>;
  undo(receipt: ChangeReceipt): Promise<void>;
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
  /** A parameterized button focuses the request field instead of inventing a value. */
  inputPrompt?: { prefix: string; message: string };
  prepare(
    selection: SelectionSnapshot,
    parameters?: FeatureParameters,
  ): { summary: string; plan: ActionPlan };
}

export interface PreparedTask {
  id: string;
  title: string;
  target: string;
  summary: string;
  changesDocument: boolean;
}

export interface TaskResult {
  message: string;
  canUndo: boolean;
}

export interface ActivityEntry {
  id: string;
  featureId: string;
  title: string;
  status: 'completed' | 'restored';
  time: number;
}
