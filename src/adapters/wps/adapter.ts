import type {
  ActionPlan,
  CellValue,
  ChangeReceipt,
  DocumentKind,
  SelectionReadOptions,
  SelectionSnapshot,
  TextActionPlan,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { isExtendedPlan } from '../../core/contracts';
import {
  assertSameSelection,
  assertSelectionLimit,
  MAX_SELECTION_CELLS,
} from '../../core/selection';
import { ReadOnlyAdapter } from '../shared/read-only-adapter';
import { getWpsApplication, type WpsApplication, type WpsDocument, type WpsRange } from './bridge';
import { validateResize } from '../shared/resize';
import { projectTextPlan, textPlanMessage, validateTextSelection } from '../../core/text-format';
import { readWpsSheetDetails } from './sheet-details';
import { readWpsPrintSettings, readWpsSheetLayout, readWpsValues, WpsTasks } from './tasks';
import {
  projectedParagraphRevision,
  readWpsParagraphs,
  restoreWpsParagraphs,
  writeWpsParagraphs,
  type WpsParagraphState,
} from './paragraphs';
import {
  isWpsNormalStyle,
  nativeWpsStyle,
  readWpsDocumentStyles,
  supportsWpsDocumentStyles,
  type WpsDocumentStyleState,
} from './document-styles';

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

function readStoryType(range: WpsRange | undefined): number | undefined {
  try {
    const storyType = range?.StoryType;
    return Number.isInteger(storyType) ? storyType : undefined;
  } catch {
    // Older bridges may not expose StoryType; existing character tools remain available.
    return undefined;
  }
}

export class WpsAdapter extends ReadOnlyAdapter {
  private readonly sessionId = crypto.randomUUID();
  private disposed = false;
  private readonly tasks: WpsTasks;
  private change?: {
    token: string;
    range: WpsRange;
    documentId: string;
    storyType?: number;
    type: TextActionPlan['type'];
    beforeSize?: number;
    beforeBold?: boolean;
    afterText: string;
    afterSize?: number;
    afterBold?: boolean;
    beforeParagraphs?: WpsParagraphState;
    afterParagraphRevision?: string;
  };
  private documentStyleChange?: {
    token: string;
    documentId: string;
    before: SelectionSnapshot;
    after: SelectionSnapshot;
    plan: Extract<TextActionPlan, { type: 'text.documentStyles' }>;
    state: WpsDocumentStyleState;
  };
  constructor(
    kind: DocumentKind,
    private readonly getApplication: () => WpsApplication = () => getWpsApplication(kind),
  ) {
    let documentStylesAvailable = false;
    if (kind === 'text') {
      try {
        documentStylesAvailable = supportsWpsDocumentStyles(getApplication());
      } catch {
        documentStylesAvailable = false;
      }
    }
    super({
      id: 'wps',
      label: kind === 'text' ? 'WPS 文字' : 'WPS 表格',
      kind,
      capabilities: [
        'selection.read',
        'document.saveCopy',
        'document.exportPdf',
        'history.undo',
        ...(kind === 'text'
          ? ([
              'text.resize',
              'text.bold',
              'paragraph.center',
              'paragraph.spacing',
              'paragraph.indent',
              'document.notice',
              ...(documentStylesAvailable ? (['text.documentStyles'] as const) : []),
            ] as const)
          : ([
              'sheet.autofit',
              'sheet.wrap',
              'sheet.header',
              'sheet.print',
              'selection.locate',
            ] as const)),
      ],
      notice:
        kind === 'text'
          ? '先选中文字。支持字号、加粗、普通正文居中、倍数行距和首行缩进；段落操作处理所在完整段落，每次最多20段，可恢复最后一次格式修改。'
          : '支持连续选区的计算、定位、列宽、换行、表头与页面设置。需要当前版本提供完整接口；打印须在 WPS 原生预览中确认。保存副本在接口完整且文件读回核验通过时报告完成；表格 PDF 的整簿导出须完整核对全部工作表（含图表表）且全部可见；局部导出仅处理当前选区或当前工作表，并须确认只选中原工作表。表格 PDF 导出均忽略原打印区域；条件不足时提供原生指引，输出后请核对实际文件。',
    });
    this.tasks = new WpsTasks({
      kind,
      application: () => this.getApplication(),
      documentId: (app) => this.documentIdentity(app),
      readSelection: (options) => this.readSelection(options),
      invalidateText: () => {
        this.change = undefined;
      },
    });
  }

  async readSelection(options?: SelectionReadOptions): Promise<SelectionSnapshot> {
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
      if (options?.includeDocumentParagraphs) {
        if (!this.info.capabilities.includes('text.documentStyles'))
          throw new HelperError('UNAVAILABLE', '当前 WPS 版本暂不支持安全读取全文标题样式。');
        const state = readWpsDocumentStyles(app, doc as WpsDocument);
        return {
          ...base,
          targetId: 'document-body',
          targetLabel: `全文（${state.paragraphs.length} 个段落）`,
          revision: JSON.stringify(state.paragraphs),
          documentParagraphs: state.paragraphs,
        };
      }
      // A collapsed WPS selection may still return the character to its right.
      const hasTextRange =
        Number.isInteger(selection.Start) &&
        Number.isInteger(selection.End) &&
        selection.Start! < selection.End!;
      const text = hasTextRange ? String(selection.Text ?? '') : '';
      const size = readSize(selection.Font?.Size);
      const bold = readBold(selection.Font);
      const range = hasTextRange ? selection.Range : undefined;
      const storyType = readStoryType(range);
      const paragraphs = hasTextRange ? readWpsParagraphs(range, app) : undefined;
      const snapshot: SelectionSnapshot = {
        ...base,
        targetId: `${selection.Start}:${selection.End}`,
        targetLabel: '当前选中文字',
        revision: JSON.stringify([text, size, bold, paragraphs?.revision, storyType]),
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
    const values: CellValue[][] = readWpsValues(selection, rows, columns);
    const address =
      typeof selection.Address === 'function' ? selection.Address() : selection.Address;
    if (!address) throw new HelperError('UNAVAILABLE', '未能确认选中的位置，请重新选择范围。');
    // Explicit layout requests keep their strict safety checks. A print-only feature must
    // still be able to return native guidance when the desktop bridge lacks those checks.
    let sheetLayout = options?.includeSheetLayout ? readWpsSheetLayout(app, selection) : undefined;
    let printSettings = options?.includePrintSettings
      ? readWpsPrintSettings(app, selection)
      : undefined;
    if (printSettings?.available && !sheetLayout) {
      try {
        sheetLayout = readWpsSheetLayout(app, selection);
      } catch (error) {
        printSettings = {
          ...printSettings,
          available: false,
          reason:
            error instanceof HelperError
              ? error.message
              : 'WPS 选区版式读取失败，无法安全设置或恢复打印参数。',
        };
      }
    }
    return {
      ...base,
      targetId: `${app.ActiveSheet?.Name ?? ''}!${address}`,
      targetLabel: address,
      revision: JSON.stringify(values),
      values,
      ...(options?.includeCellTypes ? readWpsSheetDetails(selection, app, values) : {}),
      ...(sheetLayout ? { sheetLayout } : {}),
      ...(printSettings ? { printSettings } : {}),
    };
  }

  override async apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    this.ensureActive();
    if (!isExtendedPlan(plan) && plan.type === 'text.documentStyles')
      return this.applyDocumentStyles(plan, expected);
    if (isExtendedPlan(plan)) return this.tasks.apply(plan, expected);
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
    const storyType = readStoryType(range);
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
        storyType,
      ]) !== expected.revision
    )
      throw new HelperError('STALE_SELECTION', '选中的内容变了，请重新预览。');
    this.change = undefined;
    this.documentStyleChange = undefined;
    this.tasks.clear();
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
        readStoryType(range) !== storyType ||
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
      storyType,
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
    if (this.tasks.owns(receipt)) return this.tasks.undo(receipt);
    const documentChange = this.documentStyleChange;
    if (documentChange?.token === receipt.token) {
      const app = this.getApplication();
      const doc = app.ActiveDocument;
      if (
        !doc ||
        this.documentIdentity(app) !== documentChange.documentId ||
        doc.ReadOnly ||
        (doc.ProtectionType !== undefined && doc.ProtectionType !== -1)
      )
        throw new HelperError('UNAVAILABLE', '文档已关闭、只读或受保护，无法恢复样式。');
      const current = await this.readSelection({ includeDocumentParagraphs: true });
      assertSameSelection(documentChange.after, current);
      for (const { paragraphIndex } of documentChange.plan.assignments) {
        const target = documentChange.state.targets[paragraphIndex];
        if (!target) throw new HelperError('STALE_SELECTION', '原段落已改变，请使用 WPS 的撤销。');
        target.range.Style = target.nativeStyle;
      }
      const restored = await this.readSelection({ includeDocumentParagraphs: true });
      assertSameSelection(documentChange.before, restored);
      this.documentStyleChange = undefined;
      return;
    }
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
      readStoryType(change.range) !== change.storyType ||
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
        readStoryType(change.range) !== change.storyType ||
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
    this.documentStyleChange = undefined;
    this.tasks.dispose();
  }
  async locate(selection: SelectionSnapshot, row: number, column: number): Promise<void> {
    this.ensureActive();
    await this.tasks.locate(selection, row, column);
  }
  async openResult(actionId: string): Promise<void> {
    this.ensureActive();
    await this.tasks.openResult(actionId);
  }
  private ensureActive() {
    if (this.disposed) throw new HelperError('EXPIRED_TASK', '请重新打开助手。');
  }

  private async applyDocumentStyles(
    plan: Extract<TextActionPlan, { type: 'text.documentStyles' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    if (!this.info.capabilities.includes('text.documentStyles'))
      throw new HelperError('UNAVAILABLE', '当前 WPS 版本没有可验证的内置标题样式接口。');
    validateTextSelection(plan, expected);
    assertSameSelection(expected, await this.readSelection({ includeDocumentParagraphs: true }));
    const app = this.getApplication();
    const doc = app.ActiveDocument;
    if (
      !doc ||
      this.documentIdentity(app) !== expected.documentId ||
      doc.ReadOnly ||
      (doc.ProtectionType !== undefined && doc.ProtectionType !== -1)
    )
      throw new HelperError('UNAVAILABLE', '文档已关闭、只读或受保护，无法修改全文样式。');
    const state = readWpsDocumentStyles(app, doc);
    const before: SelectionSnapshot = {
      ...this.documentIdentityFields(app, doc),
      kind: 'text',
      targetId: 'document-body',
      targetLabel: `全文（${state.paragraphs.length} 个段落）`,
      revision: JSON.stringify(state.paragraphs),
      documentParagraphs: state.paragraphs,
    };
    assertSameSelection(expected, before);
    this.change = undefined;
    this.documentStyleChange = undefined;
    this.tasks.clear();
    try {
      for (const { paragraphIndex, style } of plan.assignments) {
        const target = state.targets[paragraphIndex];
        if (!target || !isWpsNormalStyle(app, target.nativeStyle))
          throw new HelperError('STALE_SELECTION', '标题段落样式已改变，请重新预览。');
        target.range.Style = nativeWpsStyle(app, style);
      }
      const afterState = readWpsDocumentStyles(app, doc);
      const after: SelectionSnapshot = {
        ...this.documentIdentityFields(app, doc),
        kind: 'text',
        targetId: 'document-body',
        targetLabel: `全文（${afterState.paragraphs.length} 个段落）`,
        revision: JSON.stringify(afterState.paragraphs),
        documentParagraphs: afterState.paragraphs,
      };
      const projected = projectTextPlan(before, plan);
      assertSameSelection({ ...projected, revision: after.revision }, after);
      const token = `wps-document-styles:${crypto.randomUUID()}`;
      this.documentStyleChange = {
        token,
        documentId: expected.documentId,
        before,
        after,
        plan,
        state,
      };
      return {
        token,
        canUndo: true,
        message: textPlanMessage(plan),
      };
    } catch (error) {
      if (error instanceof HelperError) throw error;
      throw new HelperError(
        'INVALID_PLAN',
        '标题样式可能已部分修改，读回核对未通过。请检查全文，并使用 WPS 原生撤销。',
      );
    }
  }

  private documentIdentityFields(app: WpsApplication, doc: WpsDocument) {
    return {
      documentId: this.documentIdentity(app),
      documentName: doc.Name,
    };
  }
  private documentIdentity(app: WpsApplication) {
    const doc = this.info.kind === 'text' ? app.ActiveDocument : app.ActiveWorkbook;
    return `${this.sessionId}:${app.ActiveWindow?.Hwnd ?? ''}:${doc?.FullName || doc?.Name || ''}`;
  }
}
