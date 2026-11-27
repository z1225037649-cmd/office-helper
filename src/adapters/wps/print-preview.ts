import type {
  ChangeReceipt,
  HostAdapter,
  SelectionReadOptions,
  SelectionSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { assertSameSelection } from '../../core/selection';
import {
  getWpsApplication,
  type WpsApplication,
  type WpsWorkbook,
  type WpsWorksheet,
} from './bridge';

/**
 * Official desktop JSAPI, checked 2026-11-27 (not VBA or the web-office SDK):
 * https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Worksheet/member/PrintPreview
 * PrintPreview(false) displays the worksheet preview and disables margin/page-setup changes.
 * Its documented return type is void: a completed call is NOT proof that a window appeared.
 *
 * Integration: return withWpsPrintPreview(adapter) in the WPS branch of adapters/index.ts,
 * after the existing connection read. No task, bridge, contract, runner or UI changes required.
 * Office.js is deliberately excluded: the checked Excel Application/PageLayout references
 * provide no corresponding supported preview entry point. Do not substitute VBA/COM commands,
 * keyboard emulation, browser printing, PrintOut or shell execution.
 * https://learn.microsoft.com/en-us/javascript/api/excel/excel.application?view=excel-js-1.20
 * https://learn.microsoft.com/en-us/javascript/api/excel/excel.pagelayout?view=excel-js-1.20
 */
export interface WpsPrintPreviewWorksheet extends WpsWorksheet {
  PrintPreview?(enableChanges?: boolean): unknown;
}

const prefix = 'wps-print-preview:';
const label = '打开原生打印预览（不打印）';
const guidance =
  '请回到原 WPS 表格，选择“文件 → 打印”（或按 Ctrl+P），在原生预览中核对打印区域、纸张、分页、漏列和文字可读性；打印由你确认，助手不会提交打印。';

interface NativeTarget {
  workbook: WpsWorkbook;
  sheet: WpsPrintPreviewWorksheet;
  windowId: number | undefined;
}
interface PreviewResult extends NativeTarget {
  id: string;
  snapshot: SelectionSnapshot;
  options: SelectionReadOptions;
}

function expired(): never {
  throw new HelperError('EXPIRED_TASK', '打印预览入口已失效，请重新进行打印前检查。');
}
function stale(): never {
  throw new HelperError(
    'STALE_SELECTION',
    '原文件、工作表或打印参数已改变，请重新检查后打开预览。',
  );
}
function unavailable(): never {
  throw new HelperError('UNAVAILABLE', `当前无法安全调用原生打印预览。${guidance}`);
}

/** Keep native references within the adapter and never activate a different sheet implicitly. */
function target(app: WpsApplication): NativeTarget {
  const workbook = app.ActiveWorkbook;
  const sheet: WpsPrintPreviewWorksheet | undefined = app.ActiveSheet;
  if (!workbook || !sheet || sheet.Parent !== workbook || typeof sheet.PrintPreview !== 'function')
    unavailable();
  return { workbook, sheet, windowId: app.ActiveWindow?.Hwnd };
}

function verifyTarget(app: WpsApplication, expected: NativeTarget): void {
  // A bridge returning fresh object wrappers cannot establish this identity: fail closed.
  // Name/path equality alone must not authorize a closed/reopened, identically named book.
  if (
    app.ActiveWorkbook !== expected.workbook ||
    app.ActiveSheet !== expected.sheet ||
    expected.sheet.Parent !== expected.workbook ||
    app.ActiveWindow?.Hwnd !== expected.windowId
  )
    stale();
}

function verifySettings(snapshot: SelectionSnapshot): void {
  const settings = snapshot.printSettings;
  if (
    snapshot.kind !== 'sheet' ||
    settings?.available !== true ||
    typeof settings.printArea !== 'string' ||
    settings.printArea.replace(/\$/gu, '').toUpperCase() !==
      snapshot.targetLabel.replace(/\$/gu, '').toUpperCase() ||
    settings.orientation.toLowerCase() !== 'landscape' ||
    settings.zoom !== false ||
    settings.fitWide !== 1 ||
    settings.fitTall !== 0
  )
    stale();
}

/**
 * Opt-in adapter decoration. Only a verified, completed sheet.print receives an opaque action.
 * Preview never runs during prepare/apply/undo. Existing exports retain their own result routing.
 */
export function withWpsPrintPreview(
  host: HostAdapter,
  application: () => WpsApplication = () => getWpsApplication('sheet'),
): HostAdapter {
  if (host.info.id !== 'wps' || host.info.kind !== 'sheet') return host;
  let retained: PreviewResult | undefined;
  let disposed = false;
  let opening = false;
  const active = () => {
    if (disposed) expired();
  };

  return {
    info: host.info,
    readSelection: (options) => host.readSelection(options),
    async apply(plan, expected) {
      active();
      retained = undefined;
      const receipt = await host.apply(plan, expected);
      if (plan.type !== 'sheet.print' || (receipt.outcome && receipt.outcome !== 'completed'))
        return receipt;
      try {
        const native = target(application());
        const options: SelectionReadOptions = {
          includePrintSettings: true,
          ...(expected.sheetLayout ? { includeSheetLayout: true } : {}),
          ...(expected.cellTypes ? { includeCellTypes: true } : {}),
        };
        const snapshot = await host.readSelection(options);
        active();
        verifyTarget(application(), native);
        // Only the explicitly applied print parameters may differ from the inspected source.
        assertSameSelection({ ...expected, printSettings: snapshot.printSettings }, snapshot);
        verifySettings(snapshot);
        const id = `${prefix}${crypto.randomUUID()}`;
        retained = { ...native, id, snapshot: structuredClone(snapshot), options };
        return {
          ...receipt,
          message: `${receipt.message}可点击“${label}”请求打开 WPS 原生预览；实际分页需在原生窗口核对，未出现预览时请回到表格按 Ctrl+P。助手不会提交打印。`,
          actions: [{ id, label }, ...(receipt.actions ?? [])],
        };
      } catch {
        // Settings may already be applied. Preserve the original outcome/token/undo receipt.
        return {
          ...receipt,
          message: `${receipt.message}当前缺少可核对的安全预览入口，未提供自动预览按钮。${guidance}`,
        };
      }
    },
    async undo(receipt: ChangeReceipt) {
      active();
      retained = undefined;
      await host.undo(receipt);
    },
    ...(host.locate
      ? {
          locate: (snapshot: SelectionSnapshot, row: number, column: number) =>
            host.locate!(snapshot, row, column),
        }
      : {}),
    async openResult(id) {
      active();
      if (!id.startsWith(prefix)) {
        if (!host.openResult) expired();
        return host.openResult(id);
      }
      const result = retained;
      if (!result || result.id !== id) expired();
      if (opening) throw new HelperError('BUSY', '正在打开原生打印预览，请先查看 WPS 窗口。');
      opening = true;
      try {
        verifyTarget(application(), result);
        assertSameSelection(result.snapshot, await host.readSelection(result.options));
        active();
        if (retained !== result) expired();
        verifyTarget(application(), result);
        const preview = result.sheet.PrintPreview;
        if (typeof preview !== 'function') unavailable();
        try {
          // The sole native UI call. No activation, page-setup writes or print submission.
          await preview.call(result.sheet, false);
        } catch {
          retained = undefined;
          throw new HelperError(
            'UNAVAILABLE',
            `未能确认原生打印预览已打开，请先查看 WPS 窗口。${guidance}`,
          );
        }
      } catch (error) {
        if (error instanceof HelperError) throw error;
        retained = undefined;
        throw new HelperError('UNAVAILABLE', `未能复核原生打印预览入口。${guidance}`);
      } finally {
        opening = false;
      }
    },
    dispose() {
      retained = undefined;
      disposed = true;
      host.dispose();
    },
  };
}
