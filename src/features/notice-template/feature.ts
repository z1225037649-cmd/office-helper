import type { ExtendedActionPlan, FeatureDefinition, NoticeFields } from '../../core/contracts';
import { HelperError } from '../../core/errors';
import { noticeText, validateExtendedPlan } from '../../core/extended-actions';

const fieldNames: Record<keyof NoticeFields, string> = {
  title: '标题',
  time: '时间',
  location: '地点',
  body: '正文',
  contact: '联系人',
};

export const noticeTemplate: FeatureDefinition = {
  id: 'document.notice',
  title: '生成通知',
  description: '填写标题、时间、地点、正文和联系人，在新文档中生成通知',
  icon: 'text',
  kinds: ['text'],
  requires: ['selection.read', 'document.notice'],
  form: 'notice',
  phrases: ['生成通知', '写通知', '通知模板', '新建通知'],
  prepare(selection, parameters) {
    if (selection.kind !== 'text')
      throw new HelperError('UNAVAILABLE', '通知模板只支持文字文档，请先打开文字文档。');
    const input = parameters?.notice;
    const fields: NoticeFields = { title: '', time: '', location: '', body: '', contact: '' };
    for (const key of Object.keys(fieldNames) as (keyof NoticeFields)[]) {
      const value = input?.[key];
      if (typeof value !== 'string' || !value.trim())
        throw new HelperError('INVALID_PLAN', `请填写通知的${fieldNames[key]}。`);
      if (key !== 'body' && /[\r\n\v\f\u0085\u2028\u2029]/u.test(value))
        throw new HelperError('INVALID_PLAN', '标题、时间、地点和联系人请各填写一行。');
      fields[key] = value.trim();
    }
    const plan: ExtendedActionPlan = { type: 'document.notice', fields };
    validateExtendedPlan(plan);
    const content = noticeText(fields);
    return {
      summary: `将在新文档中生成以下完整通知，不覆盖当前文档。请核对全部内容：\n\n${content}`,
      plan,
      confirmLabel: '确认生成通知',
    };
  },
};
