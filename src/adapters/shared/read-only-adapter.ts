import type {
  ActionPlan,
  ChangeReceipt,
  HostAdapter,
  HostInfo,
  SelectionSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';

/** 原生写入与恢复经过宿主实机验证后，再逐项开放能力。 */
export abstract class ReadOnlyAdapter implements HostAdapter {
  constructor(readonly info: HostInfo) {}
  abstract readSelection(): Promise<SelectionSnapshot>;
  async apply(
    _plan: Exclude<ActionPlan, { type: 'report' }>,
    _expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    throw new HelperError('UNAVAILABLE', '当前插件框架仅接通读取，修改功能还未开放。');
  }
  async undo(_receipt: ChangeReceipt): Promise<void> {
    throw new HelperError('UNAVAILABLE', '当前插件尚未接通恢复功能。');
  }
  dispose(): void {
    /* 宿主读取不保留句柄。 */
  }
}
