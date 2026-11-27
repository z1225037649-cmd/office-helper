import type { ExtendedActionPlan, FeatureDefinition } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { validateExtendedPlan } from '../../core/extended-actions';

const printGuidance =
  '请回到原生表格，选择“文件 → 打印”（或按 Ctrl+P），核对打印区域、横向方向、宽度 1 页和高度自动；在原生打印预览中检查分页、漏列和文字可读性。压缩到一页宽可能使文字过小，需要时调整纸张或分为多页。助手不估算页数，也不会提交打印。';

export const printCheck: FeatureDefinition = {
  id: 'sheet.print',
  title: '打印前检查',
  description: '检查并确认选区打印参数；原生预览入口按宿主能力提供，打印由你确认',
  icon: 'inspect',
  kinds: ['sheet'],
  // 保留只读入口，让没有打印设置接口的宿主仍能显示原生操作指引。
  requires: ['selection.read'],
  readOptions: { includePrintSettings: true },
  phrases: ['打印前检查', '打印检查', '打印不全', '打印设置', '打印预览'],
  prepare(selection) {
    if (selection.kind !== 'sheet')
      throw new HelperError('UNAVAILABLE', '打印前检查只支持表格，请先打开表格。');
    const settings = selection.printSettings;
    if (!settings || settings.available === false) {
      const reason = settings?.reason?.trim() || '当前宿主无法自动读取或设置打印参数。';
      const message = `${reason}本次未修改打印设置。${printGuidance}`;
      return {
        summary: `为「${selection.targetLabel}」提供原生打印指引。${message}`,
        plan: { type: 'report', message, outcome: 'guidance' },
      };
    }
    if (
      settings.available !== true ||
      typeof settings.printArea !== 'string' ||
      typeof settings.orientation !== 'string' ||
      !settings.orientation.trim() ||
      !(settings.zoom === false || (Number.isFinite(settings.zoom) && settings.zoom > 0)) ||
      !Number.isSafeInteger(settings.fitWide) ||
      settings.fitWide < 0 ||
      !Number.isSafeInteger(settings.fitTall) ||
      settings.fitTall < 0 ||
      (settings.contentWidthPoints !== undefined &&
        (!Number.isFinite(settings.contentWidthPoints) || settings.contentWidthPoints < 0))
    ) {
      throw new HelperError(
        'UNAVAILABLE',
        '打印原值读取不完整，请在原生表格的“文件 → 打印”中核对。',
      );
    }
    const rows = selection.values;
    const columns = rows?.[0]?.length ?? 0;
    if (
      !Array.isArray(rows) ||
      !rows.length ||
      !columns ||
      !Array.from(rows).every((row) => Array.isArray(row) && row.length === columns)
    ) {
      throw new HelperError('EMPTY_SELECTION', '请先选择要打印的连续矩形区域。');
    }
    const orientation =
      { landscape: '横向', portrait: '纵向' }[settings.orientation.toLowerCase()] ??
      settings.orientation;
    const originalZoom = settings.zoom === false ? '按页适配' : `${settings.zoom}%`;
    const originalWide = settings.fitWide === 0 ? '自动' : `${settings.fitWide} 页`;
    const originalTall = settings.fitTall === 0 ? '自动' : `${settings.fitTall} 页`;
    const plan: ExtendedActionPlan = { type: 'sheet.print' };
    validateExtendedPlan(plan);
    return {
      summary: `「${selection.targetLabel}」打印设置原值：打印区域 ${settings.printArea || '未指定（由原生表格决定）'}；方向 ${orientation}；缩放 ${originalZoom}；宽 ${originalWide}、高 ${originalTall}。目标：当前选区为打印区域、横向、按页适配、宽度 1 页、高度自动。这里的页数参数是设置值，实际页数需在原生打印预览中核对。确认仅设置打印参数；若结果提供“打开原生打印预览（不打印）”按钮，可再点击查看，未提供时请按原生步骤操作。${printGuidance}`,
      plan,
      confirmLabel: '确认设置打印参数',
    };
  },
};
