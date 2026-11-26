import type {
  ActionPlan,
  CellValue,
  ChangeReceipt,
  DocumentKind,
  SelectionSnapshot,
  TextActionPlan,
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
import { textPlanMessage, validateTextSelection } from '../../core/text-format';
import {
  projectedParagraphRevision,
  readWpsParagraphs,
  restoreWpsParagraphs,
  writeWpsParagraphs,
  type WpsParagraphState,
} from './paragraphs';

// WPS desktop uses msoTrue (-1), msoFalse (0), and a separate mixed/unknown value.
function readBold(font: { Bold?: number | boolean } | undefined): boolean | undefined {
  try {
    const value = font?.Bold;
    if (value === true || value === -1) return true;
    if (value === false || value === 0) return false;
  } catch {
    // An older bridge may not expose Bold; resizing remains available.
  }
  return undefined;
}

function readSize(value: number | undefined): number | undefined {
  return value && value > 0 && value <= 1638 ? value : undefined;
}

export class WpsAdapter extends ReadOnlyAdapter {
  private readonly sessionId = crypto.randomUUID();
  private disposed = false;
  private change?: {
    token: string;
    range: WpsRange;
    documentId: string;
    type: TextActionPlan['type'];
    beforeSize?: number;
    beforeBold?: boolean;
    afterText: string;
    afterSize?: number;
    afterBold?: boolean;
    beforeParagraphs?: WpsParagraphState;
    afterParagraphRevision?: string;
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
        ...(kind === 'text'
          ? ([
              'text.resize',
              'text.bold',
              'paragraph.center',
              'paragraph.spacing',
              'paragraph.indent',
              'history.undo',
            ] as const)
          : []),
      ],
      notice:
        kind === 'text'
          ? '先选中文字。支持字号、加粗、普通正文居中、倍数行距和首行缩进；段落操作处理所在完整段落，每次最多20段，可恢复最后一次格式修改。'
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
      // A collapsed WPS selection may still return the character to its right.
      const hasTextRange =
        Number.isInteger(selection.Start) &&
        Number.isInteger(selection.End) &&
        selection.Start! < selection.End!;
      const text = hasTextRange ? String(selection.Text ?? '') : '';
      const size = readSize(selection.Font?.Size);
      const bold = readBold(selection.Font);
      const paragraphs = hasTextRange ? readWpsParagraphs(selection.Range, app) : undefined;
      const snapshot: SelectionSnapshot = {
        ...base,
        targetId: `${selection.Start}:${selection.End}`,
        targetLabel: '当前选中文字',
        revision: JSON.stringify([text, size, bold, paragraphs?.revision]),
        text,
        fontSize: size,
        bold,
        ...(paragraphs ? { paragraphs: paragraphs.snapshots } : {}),
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
    if (plan.type === 'text.resize') validateResize(plan, expected);
    else validateTextSelection(plan, expected);
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
    const beforeParagraphs = readWpsParagraphs(range, app);
    if (plan.type.startsWith('paragraph.') && !beforeParagraphs)
      throw new HelperError(
        'UNAVAILABLE',
        '当前版本只能处理可准确读取的普通正文段落，暂不支持表格、列表或其他文字区域。',
      );
    if (
      `${range.Start}:${range.End}` !== expected.targetId ||
      range.Text !== expected.text ||
      readSize(range.Font.Size) !== expected.fontSize ||
      readBold(range.Font) !== expected.bold ||
      JSON.stringify([
        range.Text,
        readSize(range.Font.Size),
        readBold(range.Font),
        beforeParagraphs?.revision,
      ]) !== expected.revision
    )
      throw new HelperError('STALE_SELECTION', '选中的内容变了，请重新预览。');
    this.change = undefined;
    let afterText: string;
    let afterSize: number | undefined;
    let afterBold: boolean | undefined;
    let afterParagraphs: WpsParagraphState | undefined;
    try {
      if (plan.type === 'text.resize') range.Font.Size = plan.size;
      else if (plan.type === 'text.bold') range.Font.Bold = plan.bold ? -1 : 0;
      else writeWpsParagraphs(beforeParagraphs!, plan);
      afterText = range.Text;
      afterSize = readSize(range.Font.Size);
      afterBold = readBold(range.Font);
      afterParagraphs = readWpsParagraphs(range, app);
      if (
        afterText !== expected.text ||
        afterSize !== (plan.type === 'text.resize' ? plan.size : expected.fontSize) ||
        afterBold !== (plan.type === 'text.bold' ? plan.bold : expected.bold) ||
        (beforeParagraphs && !afterParagraphs) ||
        (plan.type.startsWith('paragraph.') &&
          beforeParagraphs &&
          afterParagraphs?.revision !== projectedParagraphRevision(beforeParagraphs, plan))
      )
        throw new Error('WPS format read-back mismatch');
    } catch {
      throw new HelperError(
        'INVALID_PLAN',
        '格式可能已修改，但结果检查未完成，请查看文档并使用 WPS 的撤销。',
      );
    }
    const token = crypto.randomUUID();
    this.change = {
      token,
      range,
      documentId: expected.documentId,
      type: plan.type,
      beforeSize: expected.fontSize,
      beforeBold: expected.bold,
      afterText,
      afterSize,
      afterBold,
      beforeParagraphs,
      afterParagraphRevision: afterParagraphs?.revision,
    };
    return { token, message: textPlanMessage(plan), canUndo: true };
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
    if (
      change.range.Text !== change.afterText ||
      readSize(change.range.Font.Size) !== change.afterSize ||
      readBold(change.range.Font) !== change.afterBold ||
      readWpsParagraphs(change.range, app)?.revision !== change.afterParagraphRevision
    )
      throw new HelperError('STALE_SELECTION', '这段文字后来又被修改了，不能覆盖新的修改。');
    try {
      if (change.type === 'text.resize') change.range.Font.Size = change.beforeSize!;
      else if (change.type === 'text.bold') change.range.Font.Bold = change.beforeBold ? -1 : 0;
      else restoreWpsParagraphs(change.beforeParagraphs!, change.type);
      if (
        change.range.Text !== change.afterText ||
        readSize(change.range.Font.Size) !== change.beforeSize ||
        readBold(change.range.Font) !== change.beforeBold ||
        readWpsParagraphs(change.range, app)?.revision !== change.beforeParagraphs?.revision
      )
        throw new Error('WPS restore read-back mismatch');
    } catch {
      this.change = undefined;
      throw new HelperError('INVALID_PLAN', '恢复结果需要检查，请查看文档并使用 WPS 的撤销。');
    }
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
