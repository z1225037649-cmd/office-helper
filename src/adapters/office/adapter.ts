import type {
  ActionPlan,
  CellValue,
  ChangeReceipt,
  DocumentKind,
  HostAdapter,
  HostInfo,
  SelectionSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import {
  MAX_SELECTION_CELLS,
  assertSameSelection,
  assertSelectionLimit,
} from '../../core/selection';
import { validateResize } from '../shared/resize';

interface WordTarget {
  id: string;
  range: Word.Range;
}
interface WordChange {
  token: string;
  target: WordTarget;
  beforeSize: number;
  after: SelectionSnapshot;
}

export class OfficeAdapter implements HostAdapter {
  readonly info: HostInfo;
  private readonly sessionId = crypto.randomUUID();
  private selection?: WordTarget;
  private change?: WordChange;
  private disposed = false;

  constructor(kind: DocumentKind, canWrite = false) {
    this.info = {
      id: 'office',
      label: kind === 'text' ? 'Microsoft Word' : 'Microsoft Excel',
      kind,
      capabilities: [
        'selection.read',
        ...(canWrite && kind === 'text' ? (['text.resize', 'history.undo'] as const) : []),
      ],
      notice:
        kind === 'text'
          ? canWrite
            ? '先在文档中选中文字，再选择功能。支持恢复本次会话的最后一次字号修改。'
            : '当前 Office 版本支持读取；字号修改需要 WordApi 1.3。'
          : '选中要计算的单元格，合计会显示在助手里。',
    };
  }

  async readSelection(): Promise<SelectionSnapshot> {
    this.ensureActive();
    if (this.info.kind === 'text') {
      if (!this.info.capabilities.includes('text.resize')) {
        return Word.run((context) =>
          this.snapshot(context, {
            id: 'readonly-current-selection',
            range: context.document.getSelection(),
          }),
        );
      }
      const read = async (context: Word.RequestContext) => {
        const current = context.document.getSelection();
        let same = false;
        if (this.selection && this.info.capabilities.includes('text.resize')) {
          const comparison = current.compareLocationWith(this.selection.range);
          await context.sync();
          same = comparison.value === Word.LocationRelation.equal;
        }
        if (!same) {
          this.selection?.range.untrack();
          current.track();
          this.selection = { id: crypto.randomUUID(), range: current };
        }
        return this.snapshot(context, this.selection!);
      };
      return this.selection ? Word.run(this.selection.range, read) : Word.run(read);
    }
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      range.load(['address', 'rowCount', 'columnCount']);
      sheet.load('id');
      await context.sync();
      if (range.rowCount * range.columnCount > MAX_SELECTION_CELLS)
        throw new HelperError(
          'SELECTION_TOO_LARGE',
          '请只选中要计算的部分，最多支持 5,000 个单元格。',
        );
      range.load('values');
      await context.sync();
      return {
        ...this.documentIdentity(),
        kind: 'sheet',
        targetId: `${sheet.id}:${range.address}`,
        targetLabel: range.address,
        revision: JSON.stringify(range.values),
        values: range.values as CellValue[][],
      };
    });
  }

  async apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    this.ensureActive();
    if (!this.info.capabilities.includes('text.resize') || !this.selection)
      throw new HelperError('UNAVAILABLE', '当前版本暂不能修改文字大小。');
    validateResize(plan, expected);
    const target = this.selection;
    return Word.run(target.range, async (context) => {
      const comparison = context.document.getSelection().compareLocationWith(target.range);
      await context.sync();
      if (comparison.value !== Word.LocationRelation.equal)
        throw new HelperError('STALE_SELECTION', '选中的位置变了，请重新预览。');
      assertSameSelection(expected, await this.snapshot(context, target));
      const undoTarget = {
        id: target.id,
        range: target.range.getRange(Word.RangeLocation.whole).track(),
      };
      const previous = this.change;
      this.change = undefined;
      if (previous) this.release(previous.target.range);
      let written = false;
      try {
        target.range.font.size = plan.size;
        await context.sync();
        written = true;
        const after = await this.snapshot(context, undoTarget);
        if (after.fontSize !== plan.size || after.text !== expected.text)
          throw new HelperError('INVALID_PLAN', '文档返回的结果与预期不一致，请检查文档。');
        const token = crypto.randomUUID();
        this.change = { token, target: undoTarget, beforeSize: expected.fontSize!, after };
        return { token, message: `已把选中文字调整为 ${plan.size} 磅。`, canUndo: true };
      } catch (error) {
        undoTarget.range.untrack();
        await context.sync().catch(() => undefined);
        if (written)
          throw new HelperError(
            'INVALID_PLAN',
            '字号可能已修改，但未能完成结果检查，请查看文档并使用 Word 的撤销。',
          );
        throw error;
      }
    });
  }

  async undo(receipt: ChangeReceipt): Promise<void> {
    this.ensureActive();
    const change = this.change;
    if (!change || change.token !== receipt.token)
      throw new HelperError('EXPIRED_TASK', '这次修改已无法从助手中恢复。');
    await Word.run(change.target.range, async (context) => {
      assertSameSelection(change.after, await this.snapshot(context, change.target));
      change.target.range.font.size = change.beforeSize;
      await context.sync();
      change.target.range.font.load('size');
      await context.sync();
      if (change.target.range.font.size !== change.beforeSize)
        throw new HelperError('INVALID_PLAN', '恢复结果需要检查，请查看文档。');
      change.target.range.untrack();
      await context.sync();
    });
    this.change = undefined;
  }

  dispose(): void {
    this.disposed = true;
    if (this.selection) this.release(this.selection.range);
    if (this.change) this.release(this.change.target.range);
    this.selection = undefined;
    this.change = undefined;
  }

  private async snapshot(
    context: Word.RequestContext,
    target: WordTarget,
  ): Promise<SelectionSnapshot> {
    target.range.load('text');
    target.range.font.load('size');
    await context.sync();
    const size = target.range.font.size;
    const snapshot: SelectionSnapshot = {
      ...this.documentIdentity(),
      kind: 'text',
      targetId: target.id,
      targetLabel: '当前选中文字',
      revision: JSON.stringify([target.range.text, size]),
      text: target.range.text,
      fontSize: size > 0 ? size : undefined,
    };
    assertSelectionLimit(snapshot);
    return snapshot;
  }
  private documentIdentity() {
    const url = Office.context.document.url || '';
    let name = url.split(/[\\/]/).pop() || '当前文档';
    try {
      name = decodeURIComponent(name);
    } catch {
      /* 保留宿主返回的名称。 */
    }
    return { documentId: `${this.sessionId}:${url}`, documentName: name };
  }
  private ensureActive() {
    if (this.disposed) throw new HelperError('EXPIRED_TASK', '请重新打开助手。');
  }
  private release(range: Word.Range) {
    void Word.run(range, async (context) => {
      range.untrack();
      await context.sync();
    }).catch(() => undefined);
  }
}

export async function connectOffice(): Promise<OfficeAdapter> {
  if (typeof Office === 'undefined')
    throw new HelperError(
      'HOST_CONNECTION',
      'Office 连接组件没有加载，请从 Word 或 Excel 的加载项入口打开。',
    );
  const info = await withTimeout(Office.onReady(), 12000);
  const kind =
    info.host === Office.HostType.Word
      ? 'text'
      : info.host === Office.HostType.Excel
        ? 'sheet'
        : undefined;
  if (!kind) throw new HelperError('HOST_CONNECTION', '请从 Word 或 Excel 内打开这个助手。');
  if (!Office.context.requirements.isSetSupported(kind === 'text' ? 'WordApi' : 'ExcelApi', '1.1'))
    throw new HelperError('UNAVAILABLE', '当前 Office 版本缺少所需接口，暂时无法连接。');
  return new OfficeAdapter(
    kind,
    kind === 'text' && Office.context.requirements.isSetSupported('WordApi', '1.3'),
  );
}

export function withTimeout<T>(promise: PromiseLike<T>, delay: number): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(
      () =>
        reject(
          new HelperError('HOST_CONNECTION', '连接 Office 超时，请确认已从办公软件内打开助手。'),
        ),
      delay,
    );
    Promise.resolve(promise).then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        reject(error);
      },
    );
  });
}
