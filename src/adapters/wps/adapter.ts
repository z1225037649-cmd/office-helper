import type {
  ActionPlan,
  CellValue,
  ChangeReceipt,
  DocumentKind,
  SelectionSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import {
  assertSameSelection,
  assertSelectionLimit,
  MAX_SELECTION_CELLS,
} from '../../core/selection';
import { ReadOnlyAdapter } from '../shared/read-only-adapter';
import { getWpsApplication, type WpsApplication, type WpsRange } from './bridge';
import { validateResize } from '../shared/resize';

export class WpsAdapter extends ReadOnlyAdapter {
  private readonly sessionId = crypto.randomUUID();
  private disposed = false;
  private change?: {
    token: string;
    range: WpsRange;
    documentId: string;
    beforeSize: number;
    afterText: string;
    afterSize: number;
  };
  constructor(
    kind: DocumentKind,
    private readonly getApplication: () => WpsApplication = () => getWpsApplication(kind),
  ) {
    super({
      id: 'wps',
      label: kind === 'text' ? 'WPS 文字' : 'WPS 表格',
      kind,
      capabilities: [
        'selection.read',
        ...(kind === 'text' ? (['text.resize', 'history.undo'] as const) : []),
      ],
      notice:
        kind === 'text'
          ? '先选中想处理的文字。支持恢复本次会话的最后一次字号修改。'
          : '选中要计算的单元格，合计会显示在助手里。',
    });
  }

  async readSelection(): Promise<SelectionSnapshot> {
    this.ensureActive();
    const app = this.getApplication();
    const doc = this.info.kind === 'text' ? app.ActiveDocument : app.ActiveWorkbook;
    const selection = app.Selection;
    if (!doc || !selection)
      throw new HelperError('EMPTY_SELECTION', '请先打开文件，并选中要处理的内容。');
    const base = {
      documentId: this.documentIdentity(app),
      documentName: doc.Name,
      kind: this.info.kind,
    };
    if (this.info.kind === 'text') {
      const text = String(selection.Text ?? '');
      const size = selection.Font?.Size;
      const snapshot: SelectionSnapshot = {
        ...base,
        targetId: `${selection.Start}:${selection.End}`,
        targetLabel: '当前选中文字',
        revision: JSON.stringify([text, size]),
        text,
        fontSize: size && size > 0 && size <= 1638 ? size : undefined,
      };
      assertSelectionLimit(snapshot);
      return snapshot;
    }
    const rows = selection.Rows?.Count;
    const columns = selection.Columns?.Count;
    if (!rows || !columns) throw new HelperError('EMPTY_SELECTION', '请选中表格里的单元格。');
    if (rows * columns > MAX_SELECTION_CELLS)
      throw new HelperError(
        'SELECTION_TOO_LARGE',
        '请只选中要计算的部分，最多支持 5,000 个单元格。',
      );
    const raw = selection.Value2;
    if (raw === undefined) throw new HelperError('UNAVAILABLE', '当前 WPS 版本未能读取选区数据。');
    let values: CellValue[][];
    if (Array.isArray(raw) && raw.every(Array.isArray)) values = raw;
    else if (rows === 1 && columns === 1 && !Array.isArray(raw)) values = [[raw]];
    else throw new HelperError('UNAVAILABLE', '当前选区数据格式暂不支持，请换一段连续的单元格。');
    if (values.length !== rows || values.some((row) => row.length !== columns))
      throw new HelperError('UNAVAILABLE', '读取的单元格数量不完整，请重新选择范围。');
    const address =
      typeof selection.Address === 'function' ? selection.Address() : selection.Address;
    if (!address) throw new HelperError('UNAVAILABLE', '未能确认选中的位置，请重新选择范围。');
    return {
      ...base,
      targetId: `${app.ActiveSheet?.Name ?? ''}!${address}`,
      targetLabel: address,
      revision: JSON.stringify(values),
      values,
    };
  }

  override async apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    this.ensureActive();
    if (this.info.kind !== 'text') throw new HelperError('UNAVAILABLE', '表格暂时只支持读取。');
    validateResize(plan, expected);
    assertSameSelection(expected, await this.readSelection());
    const app = this.getApplication();
    const doc = app.ActiveDocument;
    const range = app.Selection?.Range?.Duplicate;
    if (!doc || !range)
      throw new HelperError('UNAVAILABLE', '当前 WPS 未提供稳定的文字选区，暂时无法修改。');
    if (
      this.documentIdentity(app) !== expected.documentId ||
      (range.Document.FullName || range.Document.Name) !== (doc.FullName || doc.Name)
    )
      throw new HelperError('STALE_SELECTION', '当前文件变了，请重新预览。');
    if (doc.ReadOnly || (doc.ProtectionType !== undefined && doc.ProtectionType !== -1))
      throw new HelperError('UNAVAILABLE', '文档是只读或受保护状态，请在可编辑文档中操作。');
    if (
      `${range.Start}:${range.End}` !== expected.targetId ||
      range.Text !== expected.text ||
      range.Font.Size !== expected.fontSize
    )
      throw new HelperError('STALE_SELECTION', '选中的内容变了，请重新预览。');
    this.change = undefined;
    range.Font.Size = plan.size;
    if (range.Font.Size !== plan.size || range.Text !== expected.text)
      throw new HelperError(
        'INVALID_PLAN',
        '字号可能已修改，但结果检查未完成，请查看文档并使用 WPS 的撤销。',
      );
    const token = crypto.randomUUID();
    this.change = {
      token,
      range,
      documentId: expected.documentId,
      beforeSize: expected.fontSize!,
      afterText: range.Text,
      afterSize: plan.size,
    };
    return { token, message: `已把选中文字调整为 ${plan.size} 磅。`, canUndo: true };
  }

  override async undo(receipt: ChangeReceipt): Promise<void> {
    this.ensureActive();
    const change = this.change;
    if (!change || change.token !== receipt.token)
      throw new HelperError('EXPIRED_TASK', '这次修改已无法从助手中恢复。');
    const app = this.getApplication();
    const doc = app.ActiveDocument;
    if (
      !doc ||
      this.documentIdentity(app) !== change.documentId ||
      (change.range.Document.FullName || change.range.Document.Name) !== (doc.FullName || doc.Name)
    )
      throw new HelperError('STALE_SELECTION', '当前文件变了，请回到刚才的文档再恢复。');
    if (doc.ReadOnly || (doc.ProtectionType !== undefined && doc.ProtectionType !== -1))
      throw new HelperError('UNAVAILABLE', '文档处于只读或保护状态，无法恢复。');
    if (change.range.Text !== change.afterText || change.range.Font.Size !== change.afterSize)
      throw new HelperError('STALE_SELECTION', '这段文字后来又被修改了，不能覆盖新的修改。');
    change.range.Font.Size = change.beforeSize;
    if (change.range.Font.Size !== change.beforeSize)
      throw new HelperError('INVALID_PLAN', '恢复结果需要检查，请查看文档。');
    this.change = undefined;
  }

  override dispose(): void {
    this.disposed = true;
    this.change = undefined;
  }
  private ensureActive() {
    if (this.disposed) throw new HelperError('EXPIRED_TASK', '请重新打开助手。');
  }
  private documentIdentity(app: WpsApplication) {
    const doc = this.info.kind === 'text' ? app.ActiveDocument : app.ActiveWorkbook;
    return `${this.sessionId}:${app.ActiveWindow?.Hwnd ?? ''}:${doc?.FullName || doc?.Name || ''}`;
  }
}
