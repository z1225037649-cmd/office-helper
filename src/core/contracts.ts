export type HostId = 'mock' | 'office' | 'wps';
export type DocumentKind = 'text' | 'sheet';
export type Capability = 'selection.read' | 'text.resize' | 'history.undo';
export type CellValue = string | number | boolean | null;

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
  values?: CellValue[][];
}

export type ActionPlan =
  { type: 'report'; message: string } | { type: 'text.resize'; size: number };

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
  prepare(selection: SelectionSnapshot): { summary: string; plan: ActionPlan };
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
