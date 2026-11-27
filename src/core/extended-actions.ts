import type { ExtendedActionPlan, NoticeFields } from './contracts';
import { HelperError } from './errors';

const noticeLimits: Record<keyof NoticeFields, number> = {
  title: 120,
  time: 200,
  location: 300,
  body: 12000,
  contact: 300,
};

export function validateExtendedPlan(plan: ExtendedActionPlan): void {
  if (['sheet.autofit', 'sheet.wrap', 'sheet.header', 'sheet.print'].includes(plan.type)) return;
  if (plan.type === 'document.notice') {
    for (const key of Object.keys(noticeLimits) as (keyof NoticeFields)[]) {
      const value = plan.fields?.[key];
      if (
        typeof value !== 'string' ||
        !value.trim() ||
        value.length > noticeLimits[key] ||
        /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value)
      ) {
        throw new HelperError('INVALID_PLAN', '请完整填写通知信息，并缩短过长的内容。');
      }
      if (key !== 'body' && /[\r\n]/u.test(value)) {
        throw new HelperError('INVALID_PLAN', '标题、时间、地点和联系人请各填写一行。');
      }
    }
    return;
  }
  if (plan.type === 'document.saveCopy' || plan.type === 'document.exportPdf') {
    const name = plan.fileName;
    if (
      typeof name !== 'string' ||
      !name.trim() ||
      name !== name.trim() ||
      name.length > 180 ||
      /[<>:"/\\|?*\u0000-\u001f]/u.test(name) ||
      /[. ]$/u.test(name) ||
      /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/iu.test(name) ||
      name === '.' ||
      name === '..'
    ) {
      throw new HelperError(
        'INVALID_PLAN',
        '请填写有效的新文件名，不要包含路径或 Windows 保留字符。',
      );
    }
    if (plan.type === 'document.exportPdf' && !/\.pdf$/iu.test(name)) {
      throw new HelperError('INVALID_PLAN', 'PDF 文件名需要以 .pdf 结尾。');
    }
    if (
      plan.type === 'document.exportPdf' &&
      plan.scope !== undefined &&
      !['document', 'selection', 'activeSheet'].includes(plan.scope)
    ) {
      throw new HelperError('INVALID_PLAN', '请选择有效的 PDF 导出范围。');
    }
    if (
      plan.folder !== undefined &&
      (typeof plan.folder !== 'string' ||
        plan.folder.length > 1000 ||
        /[\u0000-\u001f]/u.test(plan.folder))
    ) {
      throw new HelperError('INVALID_PLAN', '目标文件夹无效，请重新填写。');
    }
    return;
  }
  throw new HelperError('INVALID_PLAN', '这个操作暂时不能安全执行。');
}

export function noticeText(fields: NoticeFields): string {
  validateExtendedPlan({ type: 'document.notice', fields });
  return `${fields.title.trim()}\n\n时间：${fields.time.trim()}\n地点：${fields.location.trim()}\n\n${fields.body.trim()}\n\n联系人：${fields.contact.trim()}`;
}
