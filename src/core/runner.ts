import type {
  ActionPlan,
  ActivityEntry,
  ChangeReceipt,
  FeatureParameters,
  HostAdapter,
  PreparedTask,
  SelectionSnapshot,
  TaskResult,
} from './contracts';
import { HelperError } from './errors';
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

  async prepare(featureId: string, parameters?: FeatureParameters): Promise<PreparedTask> {
    return this.exclusive(async () => {
      const feature = this.registry.get(featureId);
      const reason = this.registry.unavailableReason(feature, this.host.info);
      if (reason) throw new HelperError('UNAVAILABLE', reason);
      const selection = await this.host.readSelection();
      assertSelectionLimit(selection);
      if (!feature.kinds.includes(selection.kind))
        throw new HelperError('UNAVAILABLE', '当前文件类型已改变，请重新打开助手。');
      const { summary, plan } = feature.prepare(structuredClone(selection), parameters);
      this.validatePlan(plan);
      const view: PreparedTask = Object.freeze({
        id: crypto.randomUUID(),
        title: feature.title,
        target: `${selection.documentName} · ${selection.targetLabel}`,
        summary,
        changesDocument: plan.type !== 'report',
      });
      // 单一待确认任务使旧预览无法被再次执行，也限制内存占用。
      this.pending.clear();
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
      this.validatePlan(task.plan);
      assertSameSelection(task.selection, await this.host.readSelection());
      let message: string;
      if (task.plan.type === 'report') {
        message = task.plan.message;
      } else {
        // 写入开始即作废旧恢复入口，避免部分失败时误恢复更早的操作。
        this.receipt = undefined;
        const receipt = await this.host.apply(task.plan, task.selection);
        this.receipt =
          receipt.canUndo && this.host.info.capabilities.includes('history.undo')
            ? receipt
            : undefined;
        message = receipt.message;
      }
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
      return { message, canUndo: this.canUndo };
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

  private validatePlan(plan: ActionPlan): void {
    if (plan.type === 'report' && typeof plan.message === 'string' && plan.message.trim()) return;
    validateTextPlan(plan);
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
