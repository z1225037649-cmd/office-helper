export const bidChecklistCategories = [
  'qualification',
  'technical',
  'commercial',
  'submission',
  'deadline',
  'other',
] as const;

export type BidChecklistCategory = (typeof bidChecklistCategories)[number];

export type BidChecklistStatus = 'unfilled' | 'in-progress' | 'addressed' | 'not-applicable';

export interface BidChecklistCandidate {
  readonly originalExcerpt: string;
  /** Zero-based UTF-16 offset into the input, inclusive. */
  readonly startOffset: number;
  /** Zero-based UTF-16 offset into the input, exclusive. */
  readonly endOffset: number;
  readonly category: BidChecklistCategory;
  readonly status: BidChecklistStatus;
  readonly isHeuristicCandidate: true;
}

const requirementLanguage =
  /(?:\b(?:must|shall|should|will be required|required to|is required|are required|are to|is to|have to|has to|need to|needs to|mandatory|obligatory)\b|必须|须(?:要|于|在|提交|提供|满足|符合|完成|具备|持有)?|应当|应(?:于|在|提交|提供|满足|符合|完成|具备|持有)|不得|需(?:要|于|在|提交|提供|满足|符合|完成|具备|持有))/iu;

const deadlineLanguage =
  /(?:\b(?:deadline|due date|due by|no later than|closing (?:date|time)|submission (?:deadline|date|time)|must be received by|received no later than)\b|截止(?:日期|时间|日|时)?|截至(?:日期|时间)?|递交期限|提交期限|开标时间|于.{0,32}(?:前|之前))/iu;

const imperativeLanguage =
  /^(?:(?:please|the bidder|bidders|the tenderer|tenderers|the supplier|suppliers)\s+)?(?:submit|provide|attach|upload|complete|deliver|furnish|quote|state|demonstrate|certify|sign|return|ensure|maintain|meet|supply|specify|list|confirm|enclose|present|prepare|execute|hold|possess|comply|include)\b|^(?:请)?(?:提交|提供|附上|上传|填写|递交|交付|报价|说明|证明|认证|签署|签字|盖章|确认|列明|列出|准备|执行|持有|具备|满足|遵守|包括|包含)/iu;

const listPrefix = /^(?:(?:\d+(?:\.\d+)*|[a-z]|\([a-z\d]+\))[.)]?\s+|[-*•‣]\s+)*/iu;

const categoryLanguage: ReadonlyArray<readonly [BidChecklistCategory, RegExp]> = [
  ['deadline', deadlineLanguage],
  [
    'qualification',
    /(?:\b(?:qualif(?:y|ies|ied|ication|ications)|eligib(?:le|ility)|experience|references|accreditation|certification|licen[cs]e|financial capacity|turnover|insurance|registered|registration|key personnel)\b|资格|资质|业绩|经验|信用|注册|登记|许可证|营业执照|保险|人员配置|项目负责人|财务能力|营业额)/iu,
  ],
  [
    'technical',
    /(?:\b(?:technical|specifications?|functional|performance|compatib(?:le|ility)|standards?|security|integration|architecture|service levels?|sla|warranty|support|deliverables?|implementation|methodology|solution|equipment)\b|技术|规格|参数|性能|兼容|标准|安全|集成|架构|服务水平|质保|保修|支持|交付成果|实施|方案|设备|验收)/iu,
  ],
  [
    'commercial',
    /(?:\b(?:commercial|prices?|pricing|costs?|fees?|currency|tax(?:es)?|payment|invoice|discount|budget|financial proposal|quotations?|rates?)\b|商务|价格|报价|费用|成本|货币|税费|付款|支付|发票|折扣|预算|费率)/iu,
  ],
  [
    'submission',
    /(?:\b(?:submi(?:t|ts|tted|tting|ssion)|proposals?|tenders?|bids?|uploads?|attach(?:es|ed|ing)?|attachments?|signed forms?|copies|envelopes?|portal|electronic)\b|提交|递交|投标|响应文件|投标文件|附件|签字|签章|盖章|副本|密封|电子平台|上传)/iu,
  ],
];

function isRequirementCandidate(line: string): boolean {
  const content = line.trim().replace(listPrefix, '');
  if (!content) return false;
  return (
    requirementLanguage.test(content) ||
    deadlineLanguage.test(content) ||
    imperativeLanguage.test(content)
  );
}

function categorize(line: string): BidChecklistCategory {
  const content = line.trim().replace(listPrefix, '');
  if (/^(?:please\s+)?(?:submit|upload|attach|enclose|present)\b|^(?:请)?(?:提交|上传|递交|附上)/iu.test(content))
    return 'submission';
  for (const [category, language] of categoryLanguage) {
    if (language.test(line)) return category;
  }
  return 'other';
}

/**
 * Extracts cautious, Chinese/English line-based candidates from raw tender text without changing
 * or interpreting the source. Offsets use JavaScript string indexing (UTF-16 code units); end
 * offsets are exclusive.
 */
export function analyzeBidChecklist(sourceText: string): BidChecklistCandidate[] {
  const candidates: BidChecklistCandidate[] = [];
  let lineStart = 0;

  for (let offset = 0; offset <= sourceText.length; offset += 1) {
    if (offset < sourceText.length && sourceText[offset] !== '\r' && sourceText[offset] !== '\n') {
      continue;
    }

    const originalExcerpt = sourceText.slice(lineStart, offset);
    if (isRequirementCandidate(originalExcerpt)) {
      candidates.push({
        originalExcerpt,
        startOffset: lineStart,
        endOffset: offset,
        category: categorize(originalExcerpt),
        status: 'unfilled',
        isHeuristicCandidate: true,
      });
    }

    if (offset === sourceText.length) break;
    if (sourceText[offset] === '\r' && sourceText[offset + 1] === '\n') offset += 1;
    lineStart = offset + 1;
  }

  return candidates;
}
