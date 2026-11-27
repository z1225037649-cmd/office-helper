import type {
  ExtendedActionPlan,
  FeatureDefinition,
  FeatureParameters,
  SelectionSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { validateExtendedPlan } from '../../core/extended-actions';

function prepareSave(
  type: 'document.saveCopy' | 'document.exportPdf',
  selection: SelectionSnapshot,
  parameters?: FeatureParameters,
) {
  if (selection.kind !== 'text' && selection.kind !== 'sheet')
    throw new HelperError('UNAVAILABLE', '请先打开文字文档或表格工作簿。');
  if (typeof parameters?.fileName !== 'string' || !parameters.fileName.trim()) {
    throw new HelperError(
      'INVALID_PLAN',
      type === 'document.exportPdf'
        ? '请先填写以 .pdf 结尾的目标文件名，再确认导出。'
        : '请先填写新的文件名，再确认另存副本。',
    );
  }
  const plan: ExtendedActionPlan = {
    type,
    fileName: parameters.fileName,
    ...(parameters.folder !== undefined ? { folder: parameters.folder } : {}),
    ...(type === 'document.exportPdf' && parameters.exportScope !== undefined
      ? { scope: parameters.exportScope }
      : {}),
  };
  validateExtendedPlan(plan);
  const exportScope = parameters.exportScope ?? 'document';
  if (type === 'document.saveCopy' && exportScope !== 'document')
    throw new HelperError('INVALID_PLAN', '副本保留整份原稿；需要部分内容时请选择导出 PDF。');
  if (exportScope === 'activeSheet' && selection.kind !== 'sheet')
    throw new HelperError('INVALID_PLAN', '当前工作表范围仅适用于表格。');
  if (
    exportScope === 'selection' &&
    (selection.kind === 'text' ? !selection.text?.trim() : !selection.values?.length)
  )
    throw new HelperError('EMPTY_SELECTION', '请先选中需要导出的文字或单元格，再重新预览。');
  const originalName = selection.documentName.split(/[\\/]/u).at(-1) ?? '';
  if (parameters.fileName.toLowerCase() === originalName.toLowerCase())
    throw new HelperError('INVALID_PLAN', '请填写不同于原稿的新文件名，避免同名覆盖原稿。');
  if (type === 'document.saveCopy') {
    const dot = originalName.lastIndexOf('.');
    const extension =
      dot > 0 && dot < originalName.length - 1
        ? originalName.slice(dot)
        : selection.kind === 'sheet'
          ? '.xlsx'
          : '.docx';
    const targetDot = parameters.fileName.lastIndexOf('.');
    if (
      targetDot <= 0 ||
      parameters.fileName.slice(targetDot).toLowerCase() !== extension.toLowerCase()
    )
      throw new HelperError(
        'INVALID_PLAN',
        `副本文件名必须保留原稿后缀 ${extension}，避免格式伪装或丢失宏等内容。请使用新的文件名。`,
      );
  }
  const pdf = type === 'document.exportPdf';
  const scope =
    exportScope === 'selection'
      ? `仅当前选区「${selection.targetLabel}」`
      : exportScope === 'activeSheet'
        ? '仅当前工作表（忽略原打印区域）'
        : selection.kind === 'sheet'
          ? '整个工作簿（所有工作表）'
          : '整份文档';
  const destination = parameters.folder?.trim()
    ? `目标文件夹：${parameters.folder}`
    : '目标文件夹：在原生保存对话框中选择或确认';
  return {
    summary: `将${pdf ? '导出 PDF' : '另存副本'}：${parameters.fileName}。范围：${scope}${exportScope === 'document' ? '，不以当前选区为范围' : ''}。${destination}。保留原件；如目标重名，需在原生对话框中核对并确认处理。${pdf ? 'PDF 的分页和内容完整性请在原生预览中核对；宿主无法可靠导出所选范围时提供对应原生步骤。' : ''}本次为操作预览，完成结果以宿主返回为准；取消对话框不代表${pdf ? '导出' : '另存'}成功。`,
    plan,
    confirmLabel: pdf ? '确认导出' : '确认另存',
  };
}

export const saveCopy: FeatureDefinition = {
  id: 'document.saveCopy',
  title: '另存副本',
  description: '为整份文档或整个工作簿创建副本，保留原件',
  icon: 'inspect',
  kinds: ['text', 'sheet'],
  requires: ['selection.read', 'document.saveCopy'],
  form: 'save',
  phrases: ['另存副本', '另存一份', '保存副本', '另存为'],
  prepare: (selection, parameters) => prepareSave('document.saveCopy', selection, parameters),
};

export const exportPdf: FeatureDefinition = {
  id: 'document.exportPdf',
  title: '导出 PDF',
  description: '选择整份原稿、当前选区或当前工作表导出为 PDF',
  icon: 'inspect',
  kinds: ['text', 'sheet'],
  requires: ['selection.read', 'document.exportPdf'],
  form: 'save',
  phrases: [
    '导出PDF',
    '导出 PDF',
    '保存PDF',
    '保存 PDF',
    '保存成PDF',
    '保存成 PDF',
    '转成PDF',
    '转成 PDF',
  ],
  prepare: (selection, parameters) => prepareSave('document.exportPdf', selection, parameters),
};
