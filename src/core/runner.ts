import type {
  ActionPlan,
  ActivityEntry,
  ChangeReceipt,
  FeatureParameters,
  HostAdapter,
  PreparedTask,
  ReportFinding,
  SelectionSnapshot,
  TaskResult,
} from './contracts';
import { isExtendedPlan } from './contracts';
import { HelperError } from './errors';
import { validateExtendedPlan } from './extended-actions';
import { FeatureRegistry } from './registry';
import { assertSameSelection, assertSelectionLimit } from './selection';
import { validateTextPlan } from './text-format';

interface PendingTask {
  view: PreparedTask;
  selection: SelectionSnapshot;
  plan: ActionPlan;
  featureId: string;
}

/** 界面、快捷按钮和未来的 AI 都经由同一个执行入口。 */
export class TaskRunner {
  private busy = false;
  private readonly pending = new Map<string, PendingTask>();
  private receipt?: ChangeReceipt;
  private entries: ActivityEntry[] = [];
  private report?: { selection: SelectionSnapshot; findings: ReportFinding[] };
  private resultActions = new Set<string>();

  constructor(
    readonly host: HostAdapter,
    readonly registry: FeatureRegistry,
  ) {}

  get history(): readonly ActivityEntry[] {
    return [...this.entries];
  }
  get canUndo(): boolean {
    return !!this.receipt?.canUndo;
  }

  async prepare(
    featureId: string,
    parameters?: FeatureParameters,
    expectedSelection?: SelectionSnapshot,
  ): Promise<PreparedTask> {
    return this.exclusive(async () => {
      const feature = this.registry.get(featureId);
      const reason = this.registry.unavailableReason(feature, this.host.info);
      if (reason) throw new HelperError('UNAVAILABLE', reason);
      const selection = await this.host.readSelection(feature.readOptions);
      if (expectedSelection) assertSameSelection(expectedSelection, selection);
      assertSelectionLimit(selection);
      if (!feature.kinds.includes(selection.kind))
        throw new HelperError('UNAVAILABLE', '当前文件类型已改变，请重新打开助手。');
      const { summary, plan, confirmLabel } = feature.prepare(
        structuredClone(selection),
        parameters,
      );
      this.validatePlan(plan, selection);
      const view: PreparedTask = Object.freeze({
        id: crypto.randomUUID(),
        title: feature.title,
        target: `${selection.documentName} · ${selection.targetLabel}`,
        summary,
        changesDocument: plan.type !== 'report',
        ...(confirmLabel ? { confirmLabel } : {}),
      });
      // 单一待确认任务使旧预览无法被再次执行，也限制内存占用。
      this.pending.clear();
      this.report = undefined;
      this.resultActions.clear();
      this.pending.set(view.id, {
        view,
        featureId,
        selection: structuredClone(selection),
        plan: structuredClone(plan),
      });
      return view;
    });
  }

  cancel(id: string): void {
    this.pending.delete(id);
  }

  async execute(id: string): Promise<TaskResult> {
    return this.exclusive(async () => {
      const task = this.pending.get(id);
      if (!task) throw new HelperError('EXPIRED_TASK', '这次预览已经结束，请重新选择功能。');
      this.pending.delete(id);
      const feature = this.registry.get(task.featureId);
      const reason = this.registry.unavailableReason(feature, this.host.info);
      if (reason) throw new HelperError('UNAVAILABLE', reason);
      this.validatePlan(task.plan, task.selection);
      assertSameSelection(task.selection, await this.host.readSelection(feature.readOptions));
      let message: string;
      let outcome: TaskResult['outcome'] = 'completed';
      let location: string | undefined;
      let findings: ReportFinding[] | undefined;
      let actions: TaskResult['actions'];
      if (task.plan.type === 'report') {
        message = task.plan.message;
        outcome = task.plan.outcome ?? 'completed';
        findings = task.plan.findings;
        this.report = findings?.length
          ? { selection: structuredClone(task.selection), findings: structuredClone(findings) }
          : undefined;
      } else {
        // 写入开始即作废旧恢复入口，避免部分失败时误恢复更早的操作。
        const changesSource = !task.plan.type.startsWith('document.');
        const previousReceipt = this.receipt;
        if (changesSource) this.receipt = undefined;
        const receipt = await this.host.apply(task.plan, task.selection);
        if (changesSource || receipt.canUndo) {
          this.receipt =
            receipt.canUndo && this.host.info.capabilities.includes('history.undo')
              ? receipt
              : undefined;
        }
        message = receipt.message;
        outcome = receipt.outcome ?? 'completed';
        location = receipt.location;
        actions = receipt.actions
          ?.slice(0, 3)
          .map((action) => ({ id: action.id, label: action.label }));
        this.resultActions = new Set(actions?.map((action) => action.id));
        if (changesSource && !receipt.canUndo && outcome !== 'completed')
          this.receipt = previousReceipt;
      }
      if (outcome === 'completed')
        this.entries = [
          {
            id,
            featureId: feature.id,
            title: feature.title,
            status: 'completed' as const,
            time: Date.now(),
          },
          ...this.entries,
        ].slice(0, 20);
      return { message, canUndo: this.canUndo, outcome, location, findings, actions };
    });
  }

  async openResult(actionId: string): Promise<void> {
    return this.exclusive(async () => {
      if (!this.resultActions.has(actionId) || !this.host.openResult)
        throw new HelperError('EXPIRED_TASK', '结果入口已失效，请重新导出后打开。');
      await this.host.openResult(actionId);
    });
  }

  async locate(findingId: string): Promise<void> {
    return this.exclusive(async () => {
      const report = this.report;
      const finding = report?.findings.find((item) => item.id === findingId);
      if (
        !report ||
        !finding ||
        !this.host.locate ||
        !this.host.info.capabilities.includes('selection.locate')
      ) {
        throw new HelperError('EXPIRED_TASK', '这条检查结果已失效，请重新检查后定位。');
      }
      await this.host.locate(structuredClone(report.selection), finding.row, finding.column);
      this.pending.clear();
    });
  }

  async undo(): Promise<void> {
    return this.exclusive(async () => {
      if (!this.receipt?.canUndo || !this.host.info.capabilities.includes('history.undo')) {
        throw new HelperError('UNAVAILABLE', '目前没有可以恢复的修改。');
      }
      await this.host.undo(this.receipt);
      this.receipt = undefined;
      this.pending.clear();
      this.entries = [
        {
          id: crypto.randomUUID(),
          featureId: 'history.undo',
          title: '恢复上一次修改',
          status: 'restored' as const,
          time: Date.now(),
        },
        ...this.entries,
      ].slice(0, 20);
    });
  }

  private validatePlan(plan: ActionPlan, selection: SelectionSnapshot): void {
    if (plan.type === 'report') {
      if (typeof plan.message !== 'string' || !plan.message.trim())
        throw new HelperError('INVALID_PLAN', '检查结果无效，请重新检查。');
      const ids = new Set<string>();
      if (plan.findings !== undefined) {
        if (!Array.isArray(plan.findings) || plan.findings.length > 200)
          throw new HelperError('INVALID_PLAN', '检查结果过多，请缩小范围。');
        for (const finding of plan.findings) {
          const textParagraphs = selection.documentParagraphs;
          const sheetRows = selection.values;
          const textFinding = !!textParagraphs;
          if (
            typeof finding.id !== 'string' ||
            !finding.id ||
            ids.has(finding.id) ||
            typeof finding.label !== 'string' ||
            typeof finding.description !== 'string' ||
            !Number.isInteger(finding.row) ||
            !Number.isInteger(finding.column) ||
            finding.row < 0 ||
            finding.column < 0 ||
            (textFinding
              ? finding.column !== 0 ||
                finding.row >= textParagraphs.length ||
                typeof finding.sourceExcerpt !== 'string' ||
                finding.sourceExcerpt.length > 500 ||
                typeof finding.category !== 'string' ||
                typeof finding.status !== 'string' ||
                !Number.isSafeInteger(finding.startOffset) ||
                !Number.isSafeInteger(finding.endOffset) ||
                finding.startOffset! < 0 ||
                finding.endOffset! < finding.startOffset! ||
                finding.endOffset! > (textParagraphs[finding.row]?.text.length ?? 0)
              : finding.row >= (sheetRows?.length ?? 0) ||
                finding.column >= (sheetRows?.[finding.row]?.length ?? 0))
          ) {
            throw new HelperError('INVALID_PLAN', '检查结果的位置无效，请重新检查。');
          }
          ids.add(finding.id);
        }
      }
      return;
    }
    if (isExtendedPlan(plan)) validateExtendedPlan(plan);
    else validateTextPlan(plan);
    if (!this.host.info.capabilities.includes(plan.type))
      throw new HelperError('UNAVAILABLE', '当前版本暂不能执行这个格式操作。');
  }

  private async exclusive<T>(operation: () => Promise<T>): Promise<T> {
    if (this.busy) throw new HelperError('BUSY', '正在处理上一步，请稍等。');
    this.busy = true;
    try {
      return await operation();
    } finally {
      this.busy = false;
    }
  }
}
