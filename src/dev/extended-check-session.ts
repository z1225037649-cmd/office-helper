import { createRuntime } from '../app/runtime';
import type { HostAdapter, PreparedTask, SelectionSnapshot, TaskResult } from '../core/contracts';
import { noticeText } from '../core/extended-actions';
import {
  ExtendedUnsupported,
  assertNativeFixture,
  type ExtendedNativeState,
  type ExtendedObserver,
} from './extended-check-host';
import {
  addressOnly,
  assertExtendedConfig,
  assertExtendedSelection,
  expectedFindings,
  featureParameters,
  featuresFor,
  normalizeText,
  same,
  type ExtendedConfig,
  type ExtendedFeature,
} from './extended-check-spec';

export type ExtendedStatus = 'passed' | 'unsupported' | 'manual-required' | 'manual' | 'failed';
export type ExtendedReport = (stage: string, details: Record<string, unknown>) => Promise<void>;
type Capture = { native: ExtendedNativeState; selection: SelectionSnapshot };
type Preview = { feature: ExtendedFeature; task: PreparedTask; before: Capture };
const errorCode = (error: unknown) => (error as { code?: string } | null)?.code;
function expect(value: unknown, message: string): asserts value {
  if (!value) throw new Error(message);
}

/** Development-only coordinator; all feature writes use createRuntime(host).runner. */
export class ExtendedCheckSession {
  readonly runtime;
  private claimed = false;
  private sourceId = '';
  private busy = false;
  private stopped = false;
  private preview?: Preview;
  private change?: Preview & { after: Capture };
  private records?: { before: Capture; result: TaskResult };
  private result?: TaskResult;
  readonly evidence: { check: string; status: ExtendedStatus }[] = [];

  constructor(
    readonly host: HostAdapter,
    readonly config: ExtendedConfig,
    private readonly observer: ExtendedObserver,
    private readonly report: ExtendedReport,
    private readonly local: (operation: string) => Promise<unknown>,
  ) {
    assertExtendedConfig(config);
    expect(import.meta.env.DEV, '只允许开发诊断。');
    expect(host.info.id === config.host && host.info.kind === config.kind, '宿主矩阵不匹配。');
    this.runtime = createRuntime(host);
  }

  get state() {
    return {
      claimed: this.claimed,
      stopped: this.stopped,
      pending: this.preview?.feature,
      undo: this.change?.feature,
      preview: this.preview?.task.summary,
      after: this.change?.after,
      result: this.result,
      findings: this.records?.result.findings,
    };
  }

  async initialize(): Promise<void> {
    expect(!this.claimed, '此任务窗格已领取。');
    const capture = await this.capture();
    this.sourceId = capture.selection.documentId;
    await this.report('claim', {
      run: `extended:${this.config.run}:${this.config.matrix}`,
      capabilities: this.host.info.capabilities,
    });
    this.claimed = true;
    await this.record('fixture', 'passed');
  }

  private async capture(feature?: ExtendedFeature, requireRange = true): Promise<Capture> {
    const native = await this.observer.read();
    assertNativeFixture(native, this.config);
    const featureOptions = feature ? this.runtime.registry.get(feature).readOptions : undefined;
    const selection = await this.host.readSelection({
      ...featureOptions,
      ...(this.config.kind === 'sheet' ? { includePrintSettings: true } : {}),
    });
    assertExtendedSelection(selection, this.config, requireRange);
    if (this.sourceId) expect(selection.documentId === this.sourceId, '原生文件身份已改变。');
    return { native, selection };
  }

  private async record(
    check: string,
    status: ExtendedStatus,
    details: Record<string, unknown> = {},
  ) {
    await this.report('extended-case', { check, status, ...details });
    this.evidence.push({ check, status });
  }

  async action<T>(operation: () => Promise<T>): Promise<T> {
    expect(this.claimed && !this.stopped && !this.busy, '尚未核验、已停止或仍有操作正在执行。');
    this.busy = true;
    try {
      return await operation();
    } catch (error) {
      if (
        error instanceof ExtendedUnsupported ||
        ['UNAVAILABLE', 'HOST_CONNECTION'].includes(errorCode(error) ?? '')
      ) {
        if (this.preview) this.runtime.runner.cancel(this.preview.task.id);
        this.preview = undefined;
        await this.record('operation', 'unsupported', {
          reason: error instanceof Error ? error.message : String(error),
        });
      } else {
        this.stopped = true;
        await this.record('operation', 'failed', {
          reason: error instanceof Error ? error.message : String(error),
          code: errorCode(error) ?? '',
        });
      }
      throw error;
    } finally {
      this.busy = false;
    }
  }

  async prepare(feature: ExtendedFeature): Promise<void> {
    expect(!this.change && !this.preview, '请先取消预览或完成本项恢复。');
    expect(featuresFor(this.config.kind).includes(feature), '功能不属于此矩阵。');
    const before = await this.capture(feature);
    const task = await this.runtime.runner.prepare(
      feature,
      featureParameters(feature, this.config),
    );
    this.preview = { feature, task, before };
    expect(same(before, await this.capture(feature)), '预览造成了原稿变化。');
    await this.record(`${feature}.preview`, 'passed', { summary: task.summary });
  }

  async cancel(): Promise<void> {
    const item = this.requirePreview();
    this.runtime.runner.cancel(item.task.id);
    this.preview = undefined;
    await this.rejects(() => this.runtime.runner.execute(item.task.id), 'EXPIRED_TASK');
    expect(same(item.before, await this.capture(item.feature)), '取消后原稿或选区发生变化。');
    await this.record(`${item.feature}.cancel`, 'passed', {
      scope: '原稿读回与已取消任务拒绝执行；新窗口/原生对话框另需人工核对',
    });
  }

  async supersede(): Promise<void> {
    const item = this.requirePreview();
    const replacement = await this.runtime.runner.prepare(
      item.feature,
      featureParameters(item.feature, this.config),
    );
    await this.rejects(() => this.runtime.runner.execute(item.task.id), 'EXPIRED_TASK');
    this.runtime.runner.cancel(replacement.id);
    this.preview = undefined;
    expect(same(item.before, await this.capture(item.feature)), '旧预览验证改变原稿。');
    await this.record(`${item.feature}.expired`, 'passed');
  }

  async stale(): Promise<void> {
    const item = this.requirePreview();
    const current = await this.capture(item.feature, false);
    expect(
      current.selection.targetId !== item.before.selection.targetId,
      '请先用原生 UI 移动选区，保持同一专用文件及原内容。',
    );
    expect(same(current.native, item.before.native), '移动选区时原稿内容或格式也变了，停止。');
    this.preview = undefined;
    await this.rejects(() => this.runtime.runner.execute(item.task.id), 'STALE_SELECTION');
    expect(same(current, await this.capture(item.feature, false)), '过期预览覆盖了新选区或原稿。');
    await this.record(`${item.feature}.stale`, 'passed');
  }

  async confirm(): Promise<void> {
    const item = this.requirePreview();
    expect(same(item.before, await this.capture(item.feature)), '确认前快照变化，请重新预览。');
    const file = item.feature === 'document.saveCopy' || item.feature === 'document.exportPdf';
    if (file) await this.local(item.feature.endsWith('Pdf') ? 'reserve-pdf' : 'reserve-copy');
    this.preview = undefined;
    const result = await this.runtime.runner.execute(item.task.id);
    this.result = result;
    if (item.feature === 'document.notice') {
      if (result.outcome !== 'completed') {
        expect(same(item.before, await this.capture(item.feature)), '通知指引改变了原稿。');
        await this.record(`${item.feature}.confirm`, 'unsupported', { result });
        return;
      }
      if (this.config.host === 'office') {
        // The original Office context stays bound to the source. No invented new-window API.
        expect(same(item.before.native, await this.observer.read()), '通知创建改变了原稿。');
        await this.record(`${item.feature}.confirm`, 'manual-required', {
          result,
          reason: '请在新 Word 窗口核对独立身份、未保存状态及五项完整正文。',
        });
      } else {
        const observed = await this.observer.notice();
        expect(
          observed.independent &&
            same(observed.source, item.before.native) &&
            observed.text === normalizeText(noticeText(this.config.notice)),
          '新通知独立身份、完整内容或原稿不变验证失败。',
        );
        await this.record(`${item.feature}.confirm`, 'passed', {
          result,
          sourceUnchanged: true,
          newUnsavedDocument: true,
        });
      }
      return;
    }
    const after = await this.capture(item.feature);
    if (result.outcome !== 'completed') {
      expect(same(item.before, after), '未完成操作改变了原稿。');
      await this.record(`${item.feature}.confirm`, 'unsupported', {
        outcome: result.outcome,
        result,
        reason: 'guidance/download/cancelled 不是原生功能通过；请按清单补人工证据。',
      });
      return;
    }
    if (file) {
      expect(same(item.before, after), '导出改变了原稿。');
      const expected = featureParameters(item.feature, this.config)!;
      expect(
        result.location?.replaceAll('/', '\\').toLowerCase() ===
          `${expected.folder}\\${expected.fileName}`.replaceAll('/', '\\').toLowerCase(),
        '导出位置不能核对到专用结果目录。',
      );
      const evidence = await this.local(
        item.feature.endsWith('Pdf') ? 'verify-pdf' : 'verify-copy',
      );
      await this.record(`${item.feature}.confirm`, 'manual-required', { result, evidence });
    } else if (item.feature === 'sheet.checkRecords') {
      expect(same(item.before, after), '只读检查改变了原稿。');
      expect(
        same(
          result.findings?.map(({ id, row, column }) => ({ id, row, column })),
          expectedFindings.map(({ id, row, column }) => ({ id, row, column })),
        ),
        '查重漏填条数或相对位置错误。',
      );
      this.records = { before: after, result };
      await this.record(`${item.feature}.confirm`, 'passed', { findings: result.findings });
    } else {
      expect(result.canUndo && this.runtime.runner.canUndo, '原生修改没有可恢复凭据。');
      this.change = { ...item, after };
      verifyChanged(item.feature, item.before, after);
      await this.record(`${item.feature}.confirm`, 'passed', {
        before: item.before.selection,
        after: after.selection,
      });
    }
  }

  async undo(): Promise<void> {
    const change = this.change;
    expect(change, '没有待恢复的本轮修改。');
    await this.capture(change.feature);
    await this.runtime.runner.undo();
    expect(
      same(change.before, await this.capture(change.feature)),
      '恢复后原数据、公式、逐格格式或打印原值不一致。',
    );
    this.change = undefined;
    await this.record(`${change.feature}.undo`, 'passed');
  }

  async conflict(): Promise<void> {
    const change = this.change;
    expect(change, '请先确认一项可恢复修改，再用原生 UI 修改本项属性。');
    const current = await this.capture(change.feature);
    expect(
      !same(property(change.feature, current), property(change.feature, change.after)),
      '本项属性尚无后续修改，不能声称已测冲突。',
    );
    await this.rejects(() => this.runtime.runner.undo(), 'STALE_SELECTION');
    expect(same(current, await this.capture(change.feature)), '恢复冲突覆盖了用户后续编辑。');
    await this.record(`${change.feature}.undo-conflict`, 'passed', {
      next: '按确认后快照还原冲突属性，再点恢复；不得先做下一项。',
    });
  }

  async locate(id: string): Promise<void> {
    const records = this.records;
    const expected = expectedFindings.find((finding) => finding.id === id);
    expect(records && expected, '没有本轮原始报告或未知结果。');
    const before = await this.capture('sheet.checkRecords', false);
    expect(same(before.native, records.before.native), '报告原数据已改变。');
    await this.runtime.runner.locate(id);
    const after = await this.capture('sheet.checkRecords', false);
    expect(
      same(before.native, after.native) &&
        addressOnly(after.selection.targetLabel) === expected.address,
      '定位没有落到原始报告单元格，或修改了源数据。',
    );
    await this.record(`sheet.checkRecords.locate.${id}`, 'passed', { address: expected.address });
  }

  async openResult(id: string): Promise<void> {
    expect(
      this.result?.actions?.some((action) => action.id === id),
      '未知结果操作。',
    );
    await this.capture(undefined, false);
    await this.runtime.runner.openResult(id);
    await this.record('output.open', 'manual-required', {
      reason: '原生结果入口已调用；人工确认文件及实际位置。',
    });
  }

  async verifyOutput(pdf: boolean): Promise<void> {
    await this.capture(undefined, false);
    const evidence = await this.local(pdf ? 'verify-pdf' : 'verify-copy');
    await this.record(pdf ? 'manual.pdf-file' : 'manual.copy-file', 'manual-required', {
      evidence,
    });
  }

  async manual(check: string, evidence: string): Promise<void> {
    expect(
      [
        'native-dialog-cancel',
        'native-collision-cancel',
        'native-wrong-folder-cancel',
        'print-preview',
        'output-scope-location',
        'notice-new-window',
        'switch-dedicated-file-stale',
        'records-data-stale',
      ].includes(check),
      '未知人工项。',
    );
    expect(
      evidence.trim().length >= 12 && evidence.length <= 2000,
      '请写实际软件版本、操作结果和证据位置（12–2000字）。',
    );
    await this.record(`manual.${check}`, 'manual', {
      evidence: evidence.trim(),
      acceptance: '人工补证不改变任何 unsupported/failed 自动项',
    });
  }

  async summary(): Promise<void> {
    await this.report('extended-summary', {
      acceptance: 'incomplete',
      evidence: this.evidence.slice(-60),
      evidenceCount: this.evidence.length,
      pending: this.state.pending ?? null,
      undo: this.state.undo ?? null,
      reason: '此入口不自动宣称五项四宿主总验收通过；需主线程按 CHECKLIST.md 复查所有必需项。',
    });
  }

  private requirePreview(): Preview {
    expect(this.preview, '请先选择功能并预览。');
    return this.preview;
  }
  private async rejects(operation: () => Promise<unknown>, code: string) {
    let rejected = false;
    try {
      await operation();
    } catch (error) {
      if (errorCode(error) !== code) throw error;
      rejected = true;
    }
    expect(rejected, `原生任务未按 ${code} 安全拒绝。`);
  }
}

function property(feature: ExtendedFeature, capture: Capture): unknown {
  switch (feature) {
    case 'sheet.autofit':
      return capture.native.widths;
    case 'sheet.wrap':
      return capture.native.wrap;
    case 'sheet.header':
      return capture.native.bold;
    case 'sheet.print':
      return capture.selection.printSettings;
    default:
      throw new Error('非可恢复属性。');
  }
}

function verifyChanged(feature: ExtendedFeature, before: Capture, after: Capture): void {
  const projected = structuredClone(before.native);
  if (feature === 'sheet.autofit') {
    expect(
      after.native.widths?.length === 6 &&
        after.native.widths.every((w) => Number.isFinite(w) && w > 0),
      '列宽读回不可靠。',
    );
    expect(
      after.native.widths[0] === before.native.widths?.[0] &&
        !same(after.native.widths, before.native.widths),
      '未测到列宽变化或改动了选区外列。',
    );
    projected.widths = after.native.widths;
  } else if (feature === 'sheet.wrap' || feature === 'sheet.header') {
    const key = feature === 'sheet.wrap' ? 'wrap' : 'bold';
    expect(projected[key], '缺少逐格格式读回。');
    for (let r = 2; r <= (feature === 'sheet.wrap' ? 9 : 2); r++)
      for (let c = 1; c <= 5; c++) projected[key]![r]![c] = true;
  } else if (feature === 'sheet.print') {
    const settings = after.selection.printSettings;
    expect(
      settings?.available &&
        addressOnly(settings.printArea) === 'B3:F10' &&
        settings.orientation.toLowerCase() === 'landscape' &&
        settings.zoom === false &&
        settings.fitWide === 1 &&
        settings.fitTall === 0,
      '打印参数读回不符合选区/横向/宽1页/高自动。',
    );
  } else throw new Error('未知原生写入验证。');
  expect(
    same(projected, after.native),
    '确认修改越过目标属性，或改变值、公式、顺序、隐藏行或第2表。',
  );
  if (feature !== 'sheet.print')
    expect(
      same(before.selection.printSettings, after.selection.printSettings),
      '排版修改了打印设置。',
    );
}
