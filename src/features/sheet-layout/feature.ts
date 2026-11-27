import type {
  ExtendedActionPlan,
  FeatureDefinition,
  SelectionSnapshot,
  SheetLayoutSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { validateExtendedPlan } from '../../core/extended-actions';

function requireLayout(selection: SelectionSnapshot): SheetLayoutSnapshot {
  if (selection.kind !== 'sheet')
    throw new HelperError('UNAVAILABLE', '请在表格中选择要设置格式的连续区域。');
  const values = selection.values;
  const columns = values?.[0]?.length ?? 0;
  if (!Array.isArray(values) || values.length === 0 || columns === 0)
    throw new HelperError('EMPTY_SELECTION', '请先选择要设置格式的单元格。');
  if (!Array.from(values).every((row) => Array.isArray(row) && row.length === columns))
    throw new HelperError('UNAVAILABLE', '选区不完整，请选择一个连续的矩形区域。');

  const layout = selection.sheetLayout;
  if (!layout)
    throw new HelperError('UNAVAILABLE', '无法可靠读取选区的完整布局，请在原生表格中调整格式。');
  if (layout.merged === true)
    throw new HelperError('UNAVAILABLE', '选区含合并单元格，请先在原生表格中核对后再处理。');
  if (layout.protected === true)
    throw new HelperError('UNAVAILABLE', '工作表受保护，暂不能安全修改格式。');
  if (
    layout.merged !== false ||
    layout.protected !== false ||
    !Number.isSafeInteger(layout.firstRow) ||
    layout.firstRow < 0 ||
    !Number.isSafeInteger(layout.firstColumn) ||
    layout.firstColumn < 0 ||
    typeof layout.sheetName !== 'string' ||
    !layout.sheetName.trim() ||
    !['points', 'characters'].includes(layout.widthUnit) ||
    !Array.isArray(layout.columnWidths) ||
    layout.columnWidths.length !== columns ||
    Array.from(layout.columnWidths).some((width) => !Number.isFinite(width) || width < 0) ||
    !Array.isArray(layout.headerBold) ||
    layout.headerBold.length !== columns ||
    Array.from(layout.headerBold).some((bold) => typeof bold !== 'boolean') ||
    !Array.isArray(layout.wrapText) ||
    layout.wrapText.length !== values.length ||
    Array.from(layout.wrapText).some(
      (row) =>
        !Array.isArray(row) ||
        row.length !== columns ||
        Array.from(row).some((wrapped) => typeof wrapped !== 'boolean'),
    )
  ) {
    throw new HelperError(
      'UNAVAILABLE',
      '布局或原有格式读取不完整，包含未知或混合格式，暂不能安全预览和恢复。请在原生表格中调整。',
    );
  }
  return layout;
}

function layoutPlan(type: 'sheet.autofit' | 'sheet.wrap' | 'sheet.header'): ExtendedActionPlan {
  const plan: ExtendedActionPlan = { type };
  validateExtendedPlan(plan);
  return plan;
}

export const autofitColumns: FeatureDefinition = {
  id: 'sheet.autofit',
  title: '自动调整列宽',
  description: '按内容调整选中列的整列宽度，影响选区外的同列单元格',
  icon: 'inspect',
  kinds: ['sheet'],
  requires: ['selection.read', 'sheet.autofit'],
  readOptions: { includeSheetLayout: true },
  phrases: ['自动调整列宽', '自动列宽', '调整列宽', '列宽自适应'],
  prepare(selection) {
    const layout = requireLayout(selection);
    const unit = layout.widthUnit === 'points' ? '磅' : '字符宽度单位';
    const original = layout.columnWidths
      .map((width, index) => `工作表第 ${layout.firstColumn + index + 1} 列 ${width} ${unit}`)
      .join('；');
    return {
      summary: `「${layout.sheetName}」中「${selection.targetLabel}」的列宽原值：${original}。目标：由原生表格按内容自动计算列宽，预览不估算目标数值。影响选中列的整列，包括选区外的同列单元格；列宽变化可能影响换行与分页，请确认后核对。`,
      plan: layoutPlan('sheet.autofit'),
    };
  },
};

export const wrapCells: FeatureDefinition = {
  id: 'sheet.wrap',
  title: '自动换行',
  description: '为选中单元格启用自动换行',
  icon: 'inspect',
  kinds: ['sheet'],
  requires: ['selection.read', 'sheet.wrap'],
  readOptions: { includeSheetLayout: true },
  phrases: ['自动换行', '单元格换行', '启用换行'],
  prepare(selection) {
    const layout = requireLayout(selection);
    const states = layout.wrapText.flat();
    const wrapped = states.filter(Boolean).length;
    if (wrapped === states.length)
      throw new HelperError('INVALID_PLAN', '选中单元格已经全部启用自动换行，这次无需修改。');
    return {
      summary: `只处理「${selection.targetLabel}」中的 ${states.length} 个单元格。原值：${wrapped} 格已启用、${states.length - wrapped} 格未启用自动换行。目标：全部启用自动换行。单元格内容不变，显示高度和打印分页可能改变，请在原生表格中核对。`,
      plan: layoutPlan('sheet.wrap'),
    };
  },
};

export const emphasizeHeader: FeatureDefinition = {
  id: 'sheet.header',
  title: '突出表头',
  description: '只加粗当前选区第一行',
  icon: 'inspect',
  kinds: ['sheet'],
  requires: ['selection.read', 'sheet.header'],
  readOptions: { includeSheetLayout: true },
  phrases: ['突出表头', '加粗表头', '设置表头'],
  prepare(selection) {
    const layout = requireLayout(selection);
    if (layout.headerBold.every(Boolean))
      throw new HelperError('INVALID_PLAN', '选区第一行已经全部加粗，这次无需修改。');
    const original = layout.headerBold
      .map((bold, index) => `选区第 ${index + 1} 列${bold ? '已加粗' : '未加粗'}`)
      .join('；');
    return {
      summary: `只处理「${selection.targetLabel}」的选区第一行（工作表第 ${layout.firstRow + 1} 行）。原值：${original}。目标：这一行所选单元格全部加粗；不处理选区后续行，也不把工作表第一行自动当作表头。`,
      plan: layoutPlan('sheet.header'),
    };
  },
};
