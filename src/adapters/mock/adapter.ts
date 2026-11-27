import type {
  ActionPlan,
  DocumentParagraphSnapshot,
  ChangeReceipt,
  DocumentKind,
  HostAdapter,
  HostInfo,
  SelectionReadOptions,
  SelectionSnapshot,
} from '../../core/contracts';
import { isExtendedPlan } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { noticeText, validateExtendedPlan } from '../../core/extended-actions';
import { assertSameSelection } from '../../core/selection';
import { projectTextPlan, textPlanMessage, validateTextSelection } from '../../core/text-format';

export class MockAdapter implements HostAdapter {
  readonly info: HostInfo;
  private snapshot: SelectionSnapshot;
  private restore?: {
    token: string;
    before: SelectionSnapshot;
    after: SelectionSnapshot;
    beforeDocumentParagraphs?: DocumentParagraphSnapshot[];
  };
  private documentParagraphs: DocumentParagraphSnapshot[] = [];
  readonly createdDocuments: { name: string; text: string }[] = [];
  locatedCell?: { row: number; column: number };

  constructor(kind: DocumentKind = 'text') {
    this.info = {
      id: 'mock',
      label: '演示文件',
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
              'text.documentStyles',
              'paragraph.center',
              'paragraph.spacing',
              'paragraph.indent',
              'document.notice',
            ] as const)
          : ([
              'sheet.autofit',
              'sheet.wrap',
              'sheet.header',
              'sheet.print',
              'selection.locate',
            ] as const)),
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
            bold: false,
            paragraphs: [
              {
                text: '社区活动通知',
                alignment: 'Left',
                lineSpacing: 12,
                spacingMode: 'multiple',
                firstLineIndent: 0,
                leftIndent: 0,
                rightIndent: 0,
                isListItem: false,
              },
            ],
          }
        : {
            documentId: 'demo-costs',
            documentName: '活动费用.xlsx',
            kind,
            targetId: 'costs!B2:B5',
            targetLabel: '费用表 B2:B5',
            revision: '0',
            values: [[120], [85], [45.5], [200]],
            sheetLayout: {
              columnWidths: [140],
              widthUnit: 'points',
              wrapText: [[false], [false], [false], [false]],
              headerBold: [false],
              merged: false,
              protected: false,
              firstRow: 1,
              firstColumn: 1,
              sheetName: '费用表',
            },
            printSettings: {
              available: true,
              printArea: '',
              orientation: 'Portrait',
              zoom: 100,
              fitWide: 0,
              fitTall: 0,
              contentWidthPoints: 140,
            },
          };
    if (kind === 'text')
      this.documentParagraphs = [
        { index: 0, text: '第一章 项目概况', style: 'Normal', isListItem: false, inTable: false },
        { index: 1, text: '一、项目说明', style: 'Normal', isListItem: false, inTable: false },
        { index: 2, text: '（一）服务范围', style: 'Normal', isListItem: false, inTable: false },
        {
          index: 3,
          text: '本段为示例正文，不会被标题规则改动。',
          style: 'Normal',
          isListItem: false,
          inTable: false,
        },
        { index: 4, text: '第二章 技术方案', style: 'Heading1', isListItem: false, inTable: false },
        { index: 5, text: '1. 普通列表项', style: 'Normal', isListItem: true, inTable: false },
        { index: 6, text: '第三章 表格标题', style: 'Normal', isListItem: false, inTable: true },
      ];
  }

  async readSelection(options?: SelectionReadOptions): Promise<SelectionSnapshot> {
    if (options?.includeDocumentParagraphs && this.snapshot.kind === 'text') {
      const documentParagraphs = structuredClone(this.documentParagraphs);
      return {
        documentId: this.snapshot.documentId,
        documentName: this.snapshot.documentName,
        kind: 'text',
        targetId: 'document-body',
        targetLabel: `全文（${documentParagraphs.length} 个段落）`,
        revision: JSON.stringify(documentParagraphs),
        documentParagraphs,
      };
    }
    const snapshot = structuredClone(this.snapshot);
    if (!options?.includeSheetLayout) delete snapshot.sheetLayout;
    if (!options?.includePrintSettings) delete snapshot.printSettings;
    if (snapshot.kind === 'sheet' && options?.includeCellTypes) {
      snapshot.rowScope = 'all';
      snapshot.cellTypes = snapshot.values?.map((row) =>
        row.map((value) =>
          value === null
            ? 'empty'
            : typeof value === 'number'
              ? 'number'
              : typeof value === 'boolean'
                ? 'boolean'
                : 'text',
        ),
      );
    }
    return snapshot;
  }

  async apply(
    plan: Exclude<ActionPlan, { type: 'report' }>,
    expected: SelectionSnapshot,
  ): Promise<ChangeReceipt> {
    assertSameSelection(
      expected,
      await this.readSelection({
        includeCellTypes: !!expected.cellTypes,
        includeSheetLayout: !!expected.sheetLayout,
        includePrintSettings: !!expected.printSettings,
        includeDocumentParagraphs: !!expected.documentParagraphs,
      }),
    );
    if (plan.type === 'text.documentStyles') {
      validateTextSelection(plan, expected);
      const before = structuredClone(this.documentParagraphs);
      for (const { paragraphIndex, style } of plan.assignments)
        this.documentParagraphs[paragraphIndex]!.style = style;
      const after = await this.readSelection({ includeDocumentParagraphs: true });
      const beforeSnapshot = {
        ...after,
        documentParagraphs: before,
        revision: JSON.stringify(before),
      };
      const token = crypto.randomUUID();
      this.restore = {
        token,
        before: beforeSnapshot,
        after,
        beforeDocumentParagraphs: before,
      };
      return { token, message: textPlanMessage(plan), canUndo: true };
    }
    if (isExtendedPlan(plan)) {
      validateExtendedPlan(plan);
      const token = crypto.randomUUID();
      if (!this.info.capabilities.includes(plan.type))
        throw new HelperError('UNAVAILABLE', '当前演示文件不支持这项操作。');
      if (plan.type === 'document.notice') {
        this.createdDocuments.push({
          name: `${plan.fields.title.trim()}.docx`,
          text: noticeText(plan.fields),
        });
        return {
          token,
          message: '演示：已生成一份独立的通知预览。真实宿主会创建新文档。',
          canUndo: false,
        };
      }
      if (plan.type === 'document.saveCopy' || plan.type === 'document.exportPdf') {
        return {
          token,
          message: `演示模式不会写入磁盘。真实文件需另存为“${plan.fileName}”，${plan.type === 'document.exportPdf' ? `范围选择${plan.scope === 'selection' ? '仅当前选区' : plan.scope === 'activeSheet' ? '仅当前工作表' : '整份原稿'}，` : ''}并在原软件的保存窗口选择目标文件夹、确认重名提示。`,
          canUndo: false,
          outcome: 'guidance',
        };
      }
      if (this.snapshot.kind !== 'sheet') throw new HelperError('UNAVAILABLE', '请切换到表格。');
      const before = structuredClone(this.snapshot);
      const layout = this.snapshot.sheetLayout;
      if (plan.type === 'sheet.print') {
        this.snapshot.printSettings = {
          ...this.snapshot.printSettings!,
          printArea: this.snapshot.targetId,
          orientation: 'Landscape',
          zoom: false,
          fitWide: 1,
          fitTall: 0,
        };
      } else {
        if (!layout || layout.merged || layout.protected)
          throw new HelperError('UNAVAILABLE', '此范围暂时不能安全排版。');
        if (plan.type === 'sheet.autofit')
          layout.columnWidths = layout.columnWidths.map((_, column) =>
            Math.max(
              36,
              Math.min(
                360,
                Math.max(...this.snapshot.values!.map((row) => String(row[column] ?? '').length)) *
                  9 +
                  18,
              ),
            ),
          );
        if (plan.type === 'sheet.wrap')
          layout.wrapText = layout.wrapText.map((row) => row.map(() => true));
        if (plan.type === 'sheet.header') layout.headerBold = layout.headerBold.map(() => true);
      }
      this.snapshot.revision = crypto.randomUUID();
      this.restore = { token, before, after: structuredClone(this.snapshot) };
      return {
        token,
        message:
          plan.type === 'sheet.print'
            ? '演示：已设置选区打印、横向、宽度一页和高度自动。真实打印仍须在原软件按 Ctrl+P 检查列是否完整、缩放后是否清楚，再提交打印。'
            : '已在演示表格应用排版，可以恢复。',
        canUndo: true,
      };
    }
    validateTextSelection(plan, expected);
    const before = structuredClone(this.snapshot);
    const token = crypto.randomUUID();
    this.snapshot = projectTextPlan(this.snapshot, plan);
    this.snapshot.revision = crypto.randomUUID();
    this.restore = { token, before, after: structuredClone(this.snapshot) };
    return { token, message: textPlanMessage(plan), canUndo: true };
  }

  async locate(selection: SelectionSnapshot, row: number, column: number): Promise<void> {
    assertSameSelection(
      selection,
      await this.readSelection({ includeCellTypes: !!selection.cellTypes }),
    );
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(column) ||
      row < 0 ||
      column < 0 ||
      row >= (selection.values?.length ?? 0) ||
      column >= (selection.values?.[row]?.length ?? 0)
    ) {
      throw new HelperError('INVALID_PLAN', '定位位置无效。');
    }
    this.locatedCell = { row, column };
  }

  async undo(receipt: ChangeReceipt): Promise<void> {
    if (!this.restore || receipt.token !== this.restore.token)
      throw new HelperError('EXPIRED_TASK', '这次修改已无法恢复。');
    if (this.restore.after.documentParagraphs) {
      assertSameSelection(
        this.restore.after,
        await this.readSelection({ includeDocumentParagraphs: true }),
      );
      this.documentParagraphs = structuredClone(this.restore.beforeDocumentParagraphs ?? []);
      this.restore = undefined;
      return;
    }
    assertSameSelection(this.restore.after, this.snapshot);
    this.snapshot = { ...this.restore.before, revision: crypto.randomUUID() };
    this.restore = undefined;
  }

  dispose(): void {
    this.restore = undefined;
  }
}
