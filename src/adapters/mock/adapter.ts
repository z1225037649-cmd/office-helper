import type {
  ActionPlan,
  ChangeReceipt,
  DocumentKind,
  HostAdapter,
  HostInfo,
  SelectionSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { assertSameSelection } from '../../core/selection';

export class MockAdapter implements HostAdapter {
  readonly info: HostInfo;
  private snapshot: SelectionSnapshot;
  private restore?: { token: string; before: SelectionSnapshot; after: SelectionSnapshot };

  constructor(kind: DocumentKind = 'text') {
    this.info = {
      id: 'mock',
      label: '演示文件',
      kind,
      capabilities: [
        'selection.read',
        ...(kind === 'text' ? (['text.resize', 'history.undo'] as const) : []),
      ],
    };
    this.snapshot =
      kind === 'text'
        ? {
            documentId: 'demo-notice',
            documentName: '社区活动通知.docx',
            kind,
            targetId: 'title',
            targetLabel: '通知标题',
            revision: '0',
            text: '社区活动通知',
            fontSize: 18,
          }
        : {
            documentId: 'demo-costs',
            documentName: '活动费用.xlsx',
            kind,
            targetId: 'costs!B2:B5',
            targetLabel: '费用表 B2:B5',
            revision: '0',
            values: [[120], [85], [45.5], [200]],
          };
  }

  async readSelection(): Promise<SelectionSnapshot> {
    return structuredClone(this.snapshot);
  }

  async apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    assertSameSelection(expected, this.snapshot);
    if (
      this.snapshot.kind !== 'text' ||
      plan.type !== 'text.resize' ||
      !Number.isFinite(plan.size) ||
      plan.size < 8 ||
      plan.size > 72
    ) {
      throw new HelperError('INVALID_PLAN', '这个文件暂不支持此操作。');
    }
    const before = structuredClone(this.snapshot);
    const token = crypto.randomUUID();
    this.snapshot.fontSize = plan.size;
    this.snapshot.revision = crypto.randomUUID();
    this.restore = { token, before, after: structuredClone(this.snapshot) };
    return { token, message: `已把选中文字放大到 ${plan.size} 磅。`, canUndo: true };
  }

  async undo(receipt: ChangeReceipt): Promise<void> {
    if (!this.restore || receipt.token !== this.restore.token)
      throw new HelperError('EXPIRED_TASK', '这次修改已无法恢复。');
    assertSameSelection(this.restore.after, this.snapshot);
    this.snapshot = { ...this.restore.before, revision: crypto.randomUUID() };
    this.restore = undefined;
  }

  dispose(): void {
    this.restore = undefined;
  }
}
