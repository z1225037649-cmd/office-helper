import type {
  ActionPlan,
  BuiltInParagraphStyle,
  CellType,
  CellValue,
  ChangeReceipt,
  DocumentKind,
  HostAdapter,
  HostInfo,
  SelectionReadOptions,
  SelectionSnapshot,
  TextActionPlan,
} from '../../core/contracts';
import { isExtendedPlan } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import {
  MAX_SELECTION_CELLS,
  MAX_DOCUMENT_PARAGRAPHS,
  MAX_DOCUMENT_TEXT,
  MAX_SELECTION_PARAGRAPHS,
  assertSameSelection,
  assertSelectionLimit,
} from '../../core/selection';
import { projectTextPlan, textPlanMessage, validateTextSelection } from '../../core/text-format';
import { explicitMultipleSpacing } from './spacing';
import { OfficeTasks, officeTaskCapabilities } from './tasks';

interface WordTarget {
  id: string;
  range: Word.Range;
}
interface WordChange {
  token: string;
  target: WordTarget;
  before: SelectionSnapshot;
  plan: Exclude<TextActionPlan, { type: 'text.documentStyles' }>;
  after: SelectionSnapshot;
}
interface WordDocumentStyleChange {
  token: string;
  before: SelectionSnapshot;
  after: SelectionSnapshot;
  plan: Extract<TextActionPlan, { type: 'text.documentStyles' }>;
}
type WordChangeState = WordChange | WordDocumentStyleChange;

function builtInParagraphStyle(style: string): BuiltInParagraphStyle {
  const value = style.replace(/[\s_-]/gu, '').toLocaleLowerCase();
  if (value === 'normal') return 'Normal';
  if (value === 'title') return 'Title';
  if (value === 'subtitle') return 'Subtitle';
  if (value === 'heading1') return 'Heading1';
  if (value === 'heading2') return 'Heading2';
  if (value === 'heading3') return 'Heading3';
  return 'Other';
}

function readCellTypes(value: unknown): CellType[][] | undefined {
  if (!Array.isArray(value) || value.some((row) => !Array.isArray(row))) return undefined;
  return value.map((row: unknown[]) =>
    row.map((type): CellType => {
      switch (type) {
        case 'Empty':
          return 'empty';
        case 'String':
          return 'text';
        case 'Integer':
        case 'Double':
          return 'number';
        case 'Boolean':
          return 'boolean';
        case 'Error':
          return 'error';
        default:
          return 'unknown';
      }
    }),
  );
}

export class OfficeAdapter implements HostAdapter {
  readonly info: HostInfo;
  private readonly sessionId = crypto.randomUUID();
  private selection?: WordTarget;
  private change?: WordChangeState;
  private disposed = false;
  private readonly tasks: OfficeTasks;

  constructor(kind: DocumentKind, canWrite = false) {
    this.info = {
      id: 'office',
      label: kind === 'text' ? 'Microsoft Word' : 'Microsoft Excel',
      kind,
      capabilities: [
        'selection.read',
        ...(canWrite && kind === 'text'
          ? ([
              'text.resize',
              'text.bold',
              'text.documentStyles',
              'paragraph.center',
              'paragraph.spacing',
              'paragraph.indent',
              'history.undo',
            ] as const)
          : []),
        ...officeTaskCapabilities(kind),
      ],
      notice:
        kind === 'text'
          ? canWrite
            ? '先在文档中选中文字。段落功能会处理所在的完整段落，每次最多20段；可恢复本次会话的最后一次格式修改。'
            : '当前 Office 版本支持读取；文字格式修改需要 WordApi 1.3。'
          : '选中要计算的单元格，合计或平均值会显示在助手里。',
    };
    this.tasks = new OfficeTasks(
      kind,
      () => this.documentIdentity(),
      async (expected) => {
        assertSameSelection(expected, await this.readSelection());
      },
    );
  }

  async readSelection(options: SelectionReadOptions = {}): Promise<SelectionSnapshot> {
    this.ensureActive();
    if (this.info.kind === 'text') {
      if (options.includeDocumentParagraphs) {
        if (!this.info.capabilities.includes('text.documentStyles'))
          throw new HelperError('UNAVAILABLE', '当前 Word 版本暂不支持安全读取全文标题样式。');
        return Word.run((context) => this.documentStyleSnapshot(context));
      }
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
      range.load(options.includeCellTypes ? ['values', 'valueTypes'] : 'values');
      await context.sync();
      const snapshot: SelectionSnapshot = {
        ...this.documentIdentity(),
        kind: 'sheet',
        targetId: `${sheet.id}:${range.address}`,
        targetLabel: range.address,
        revision: JSON.stringify(range.values),
        values: range.values as CellValue[][],
        ...(options.includeCellTypes
          ? { cellTypes: readCellTypes(range.valueTypes), rowScope: 'all' as const }
          : {}),
      };
      return this.tasks.capture(context, range, sheet.id, snapshot, options);
    });
  }

  async apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    this.ensureActive();
    if (isExtendedPlan(plan)) {
      if (!this.info.capabilities.includes(plan.type))
        throw new HelperError('UNAVAILABLE', '当前 Office 版本不支持这个操作，请使用原生界面。');
      return this.tasks.apply(plan, expected);
    }
    if (plan.type === 'text.documentStyles') return this.applyDocumentStyles(plan, expected);
    if (!this.info.capabilities.includes(plan.type) || !this.selection)
      throw new HelperError('UNAVAILABLE', '当前版本暂不能执行这个格式操作。');
    validateTextSelection(plan, expected);
    const target = this.selection;
    return Word.run(target.range, async (context) => {
      const comparison = context.document.getSelection().compareLocationWith(target.range);
      await context.sync();
      if (comparison.value !== Word.LocationRelation.equal)
        throw new HelperError('STALE_SELECTION', '选中的位置变了，请重新预览。');
      const state = await this.readTextState(context, target);
      assertSameSelection(expected, state.snapshot);
      const undoTarget = {
        id: target.id,
        range: target.range.getRange(Word.RangeLocation.whole).track(),
      };
      const previous = this.change;
      this.change = undefined;
      if (previous) this.releaseChange(previous);
      let written = false;
      try {
        written = true;
        this.writeFormat(target.range, state.paragraphs, plan);
        await context.sync();
        const after = await this.snapshot(context, undoTarget);
        assertSameSelection(
          { ...projectTextPlan(expected, plan), revision: after.revision },
          after,
        );
        const token = crypto.randomUUID();
        this.change = { token, target: undoTarget, before: expected, plan, after };
        return { token, message: textPlanMessage(plan), canUndo: true };
      } catch (error) {
        undoTarget.range.untrack();
        await context.sync().catch(() => undefined);
        if (written)
          throw new HelperError(
            'INVALID_PLAN',
            '格式可能已修改，但未能完成结果检查，请查看文档并使用 Word 的撤销。',
          );
        throw error;
      }
    });
  }

  async undo(receipt: ChangeReceipt): Promise<void> {
    this.ensureActive();
    if (this.tasks.owns(receipt)) return this.tasks.undo(receipt);
    const change = this.change;
    if (!change || change.token !== receipt.token)
      throw new HelperError('EXPIRED_TASK', '这次修改已无法从助手中恢复。');
    if (!('target' in change)) {
      await Word.run(async (context) => {
        const current = await this.documentStyleSnapshot(context);
        assertSameSelection(change.after, current);
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        for (const { paragraphIndex } of change.plan.assignments) {
          const before = change.before.documentParagraphs?.[paragraphIndex];
          const paragraph = paragraphs.items[paragraphIndex];
          if (!before || !paragraph || before.style !== 'Normal')
            throw new HelperError('STALE_SELECTION', '原段落样式无法准确恢复，请使用 Word 的撤销。');
          paragraph.styleBuiltIn = before.style;
        }
        await context.sync();
        const restored = await this.documentStyleSnapshot(context);
        assertSameSelection(change.before, restored);
      });
      this.change = undefined;
      return;
    }
    await Word.run(change.target.range, async (context) => {
      const state = await this.readTextState(context, change.target);
      assertSameSelection(change.after, state.snapshot);
      this.writeFormat(change.target.range, state.paragraphs, change.plan, change.before);
      await context.sync();
      const restored = await this.snapshot(context, change.target);
      assertSameSelection({ ...change.before, revision: restored.revision }, restored);
      change.target.range.untrack();
      await context.sync();
    });
    this.change = undefined;
  }

  dispose(): void {
    this.disposed = true;
    this.tasks.dispose();
    if (this.selection) this.release(this.selection.range);
    if (this.change) this.releaseChange(this.change);
    this.selection = undefined;
    this.change = undefined;
  }

  async locate(snapshot: SelectionSnapshot, row: number, column: number): Promise<void> {
    this.ensureActive();
    if (this.info.kind !== 'sheet')
      throw new HelperError('UNAVAILABLE', '请在原始 Excel 表格中定位检查结果。');
    return this.tasks.locate(snapshot, row, column);
  }

  async openResult(actionId: string): Promise<void> {
    this.ensureActive();
    return this.tasks.openResult(actionId);
  }

  private async snapshot(
    context: Word.RequestContext,
    target: WordTarget,
  ): Promise<SelectionSnapshot> {
    return (await this.readTextState(context, target)).snapshot;
  }

  private async documentStyleSnapshot(context: Word.RequestContext): Promise<SelectionSnapshot> {
    const collection = context.document.body.paragraphs;
    collection.load('items');
    await context.sync();
    if (collection.items.length > MAX_DOCUMENT_PARAGRAPHS)
      throw new HelperError('SELECTION_TOO_LARGE', '全文段落超出本次安全检查上限。');
    for (const paragraph of collection.items)
      paragraph.load(['text', 'styleBuiltIn', 'isListItem', 'tableNestingLevel']);
    await context.sync();
    let textLength = 0;
    const documentParagraphs = collection.items.map((paragraph, index) => {
      const text = paragraph.text;
      textLength += text.length;
      return {
        index,
        text,
        style: builtInParagraphStyle(paragraph.styleBuiltIn),
        isListItem: paragraph.isListItem,
        inTable: paragraph.tableNestingLevel > 0,
      };
    });
    if (textLength > MAX_DOCUMENT_TEXT)
      throw new HelperError('SELECTION_TOO_LARGE', '全文字符数超出本次安全检查上限。');
    const snapshot: SelectionSnapshot = {
      ...this.documentIdentity(),
      kind: 'text',
      targetId: 'document-body',
      targetLabel: `全文（${documentParagraphs.length} 个段落）`,
      revision: JSON.stringify(documentParagraphs),
      documentParagraphs,
    };
    assertSelectionLimit(snapshot);
    return snapshot;
  }

  private async applyDocumentStyles(
    plan: Extract<TextActionPlan, { type: 'text.documentStyles' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    if (!this.info.capabilities.includes('text.documentStyles'))
      throw new HelperError('UNAVAILABLE', '当前 Word 版本暂不支持全文标题样式。');
    validateTextSelection(plan, expected);
    return Word.run(async (context) => {
      const before = await this.documentStyleSnapshot(context);
      assertSameSelection(expected, before);
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      const previous = this.change;
      this.change = undefined;
      if (previous) this.releaseChange(previous);
      try {
        for (const { paragraphIndex, style } of plan.assignments)
          paragraphs.items[paragraphIndex]!.styleBuiltIn = style;
        await context.sync();
        const after = await this.documentStyleSnapshot(context);
        const projected = projectTextPlan(before, plan);
        assertSameSelection({ ...projected, revision: after.revision }, after);
        const token = crypto.randomUUID();
        this.change = { token, before, after, plan };
        return {
          token,
          message: textPlanMessage(plan),
          canUndo: true,
        };
      } catch (error) {
        throw new HelperError(
          'INVALID_PLAN',
          '标题样式可能已部分修改，读回核对未通过。请检查全文，并使用 Word 原生撤销。',
        );
      }
    });
  }

  private async readTextState(context: Word.RequestContext, target: WordTarget) {
    target.range.load('text');
    target.range.font.load('size');
    const writable = this.info.capabilities.includes('text.resize');
    if (writable) target.range.font.load('bold');
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
    let paragraphs: Word.Paragraph[] = [];
    if (writable) {
      snapshot.bold =
        typeof target.range.font.bold === 'boolean' ? target.range.font.bold : undefined;
      const collection = target.range.paragraphs;
      collection.load('items');
      await context.sync();
      if (collection.items.length > MAX_SELECTION_PARAGRAPHS)
        throw new HelperError('SELECTION_TOO_LARGE', '请缩小范围，每次最多处理20个完整段落。');
      paragraphs = collection.items;
      for (const p of paragraphs)
        p.load([
          'text',
          'alignment',
          'lineSpacing',
          'firstLineIndent',
          'leftIndent',
          'rightIndent',
          'isListItem',
        ]);
      const xml = target.range.getOoxml();
      await context.sync();
      const spacing = explicitMultipleSpacing(
        xml.value,
        paragraphs.map((p) => p.lineSpacing),
      );
      snapshot.paragraphs = paragraphs.map((p, i) => ({
        text: p.text,
        alignment: p.alignment,
        lineSpacing: p.lineSpacing,
        spacingMode: spacing[i] ? ('multiple' as const) : ('unsupported' as const),
        firstLineIndent: p.firstLineIndent,
        leftIndent: p.leftIndent,
        rightIndent: p.rightIndent,
        isListItem: p.isListItem,
      }));
      snapshot.revision = JSON.stringify([
        snapshot.text,
        snapshot.fontSize,
        snapshot.bold,
        snapshot.paragraphs,
      ]);
      assertSelectionLimit(snapshot);
    }
    return { snapshot, paragraphs };
  }

  /** Only the requested native property is written, including during recovery. */
  private writeFormat(
    range: Word.Range,
    paragraphs: Word.Paragraph[],
    plan: TextActionPlan,
    restore?: SelectionSnapshot,
  ) {
    if (plan.type === 'text.resize') range.font.size = restore?.fontSize ?? plan.size;
    else if (plan.type === 'text.bold') range.font.bold = restore?.bold ?? plan.bold;
    else
      paragraphs.forEach((paragraph, index) => {
        const before = restore?.paragraphs?.[index];
        if (plan.type === 'paragraph.center')
          paragraph.alignment = (before?.alignment ?? 'Centered') as Word.Alignment;
        else if (plan.type === 'paragraph.spacing')
          paragraph.lineSpacing = before?.lineSpacing ?? plan.multiple * 12;
        else if (plan.type === 'paragraph.indent')
          paragraph.firstLineIndent = before?.firstLineIndent ?? plan.points;
      });
  }
  private documentIdentity() {
    const url = Office.context.document.url || '';
    let path = url;
    try {
      const parsed = new URL(url);
      if (['http:', 'https:', 'file:'].includes(parsed.protocol)) path = parsed.pathname;
    } catch {
      /* Native paths and unsaved documents need not be URLs. */
    }
    let name = path.split(/[\\/]/).pop() || '当前文档';
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
  private releaseChange(change: WordChangeState) {
    if ('target' in change) this.release(change.target.range);
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
