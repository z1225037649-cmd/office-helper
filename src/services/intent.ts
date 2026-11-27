import type { FeatureDefinition, FeatureParameters } from '../core/contracts';

export type IntentResult =
  | { type: 'match'; featureId: string; parameters?: FeatureParameters }
  | { type: 'choices'; featureIds: string[] }
  | { type: 'clarify'; message: string }
  | { type: 'unknown' };

/** 未来远程 AI 只返回功能 ID；不得直接执行模型生成的代码。 */
export interface IntentProvider {
  resolve(
    input: string,
    available: readonly FeatureDefinition[],
    signal?: AbortSignal,
  ): Promise<IntentResult>;
}

const numericPattern = /[+-]?(?:\d+(?:\.\d*)?|\.\d+)/g;
const negatedRequest =
  /不要|不(?:用|必|想|需要|予|再|宜|能|得|准|应|可|加粗|居中|缩进|放大|改|调|设)|无需|无须|别|勿|避免|禁止|取消|撤销|停止/;
const multipleSteps = /然后|接着|之后|同时|并且|以及|顺便|再|并|和|或者|或|且|后/;

/** Only a complete supported request can become a mutation; unknown remainder needs clarification. */
function wholeRequest(input: string, operation: string) {
  const prefix = '(?:请|你|帮我|帮忙|麻烦|把|将|给我|给|让|先)*';
  const target =
    '(?:选中|选区|当前|这些|这段|这一段|这部分|本段|文字|字体|字|文本|段落|标题|内容|它|的)*';
  return new RegExp(`^${prefix}${target}(?:${operation})${target}(?:一下|显示)?$`).test(input);
}

function clarify(message: string): IntentResult {
  return { type: 'clarify', message };
}

function numberIn(input: string) {
  const numbers = [...input.matchAll(numericPattern)];
  if (numbers.length !== 1 || !numbers[0] || numbers[0].index === undefined) return undefined;
  const match = numbers[0];
  // A signed value can mean a relative change. Require an unambiguous absolute target.
  if (/^[+-]/.test(match[0])) return undefined;
  const value = Number(match[0]);
  if (!Number.isFinite(value)) return undefined;
  return {
    value,
    before: input.slice(0, match.index),
    after: input.slice(match.index + match[0].length),
  };
}

function fontSizeIntent(input: string, featureId: string): IntentResult {
  const number = numberIn(input);
  if (number) {
    const target =
      '放大(?:到|为|成)?|调整到|调到|改成|改为|设置为|设为|设到|字号(?:调整到|调到|改为|设为)?|到';
    if (!wholeRequest(number.before, target) || !/^(?:磅|点|pt)?$/.test(number.after))
      return clarify('请说一个明确的目标字号，使用“磅”，例如“把字放大到 30 磅”。');
    if (number.value < 8 || number.value > 72) return clarify('字号请填写 8 到 72 磅之间的数字。');
    if (!Number.isInteger(number.value * 2))
      return clarify('本版字号支持0.5磅的精度，例如“把字放大到30.5磅”。');
    return { type: 'match', featureId, parameters: { fontSize: number.value } };
  }
  // An explicit, incomplete, or unsupported value must never fall back to an increase.
  if (
    /\d|放大到|调整到|调到|改成|改为|设置为|设为|设到|字号|到|[二两三四五六七八九十百零]点|(?:磅|号|pt|厘米|毫米|英寸|cm|mm|px|倍)/.test(
      input,
    )
  )
    return clarify('请告诉我一个 8 到 72 磅之间的具体字号，例如“把字放大到 30 磅”。');
  if (!wholeRequest(input, '(?:放大(?:一点|一些)?|大一点|看不清)'))
    return clarify('请单独说明字号操作，例如“把字放大一点”或“把字放大到 30 磅”。');
  return { type: 'match', featureId };
}

function spacingIntent(input: string, featureId: string): IntentResult {
  const number = numberIn(input);
  if (
    !number ||
    /固定|最小|至少|最大|最少/.test(input) ||
    !/^(?:倍|倍行距)$/.test(number.after) ||
    !(
      wholeRequest(number.before, '行距(?:(?:调整|设置|调|改|设)(?:到|为|成)?|到|为|成)?') ||
      (number.after === '倍行距' &&
        (wholeRequest(number.before, '调整到|调到|改成|改为|设置为|设为|设到') ||
          /^(?:请|帮我|麻烦)*$/.test(number.before)))
    ) ||
    number.value < 1 ||
    number.value > 3 ||
    Math.abs(number.value * 20 - Math.round(number.value * 20)) > 1e-8
  )
    return clarify(
      '请明确说 1 到 3 倍行距，支持0.05倍的精度，例如“行距设为 1.5 倍”；暂不支持固定磅值行距。',
    );
  return {
    type: 'match',
    featureId,
    parameters: { lineSpacingMultiple: number.value },
  };
}

function indentIntent(input: string, featureId: string): IntentResult {
  const number = numberIn(input);
  if (
    !number ||
    /悬挂|左缩进|右缩进/.test(input) ||
    !wholeRequest(number.before, '首行缩进(?:(?:调整|设置|调|改|设)(?:到|为|成)?|到|为|成)?')
  )
    return clarify('请明确首行缩进的数值和单位，例如“首行缩进 24 磅”或“首行缩进 0.74 厘米”。');
  if (!/^(?:磅|pt|厘米|cm)$/.test(number.after))
    return clarify('首行缩进请使用磅或厘米，例如“首行缩进 24 磅”；暂不支持按字符缩进。');
  const points = /^(?:厘米|cm)$/.test(number.after) ? (number.value * 72) / 2.54 : number.value;
  if (!Number.isFinite(points) || points < 0 || points > 144 + 1e-9)
    return clarify('首行缩进请填写 0 到 144 磅之间的数值，或不超过 5.08 厘米。');
  return {
    type: 'match',
    featureId,
    parameters: { firstLineIndentPoints: Math.round(points * 20) / 20 },
  };
}

function averageIntent(input: string, featureId: string): IntentResult {
  let request = input;
  // Consume a bounded number of courtesy tokens and exactly one target. Repeated,
  // overlapping target alternatives must never enter an unbounded backtracking regex.
  for (let i = 0; i < 4; i++) {
    const prefix = request.match(/^(?:请|你|帮我|帮忙|麻烦|把|将|给我|给|先)/)?.[0];
    if (!prefix) break;
    request = request.slice(prefix.length);
  }
  const operationPattern = /^(?:计算一下|查看一下|算一下|求一下|看看一下|计算|查看|看看|算|求)/;
  const targetPattern =
    /^(?:(?:当前|所)?选中(?:的)?(?:范围|单元格|数字|数据|金额)?|选区|当前范围|这些(?:单元格|数字|数据|金额)?|单元格|范围|数字|数据|金额)(?:的)?/;
  let operation = request.match(operationPattern)?.[0];
  if (operation) request = request.slice(operation.length);
  const target = request.match(targetPattern)?.[0];
  if (target) request = request.slice(target.length);
  if (!operation) {
    operation = request.match(operationPattern)?.[0];
    if (operation) request = request.slice(operation.length);
  }
  if (
    !/^(?:平均值|平均数)(?:一下|显示)?$/.test(request) &&
    !(operation === '求' && /^平均(?:一下|显示)?$/.test(request))
  )
    return clarify(
      '请单独说“算一下选中范围的平均值”。本版包含隐藏和筛选行，不支持只算可见行、指定结果数值或写回单元格。',
    );
  return { type: 'match', featureId };
}

export class LocalIntentProvider implements IntentProvider {
  async resolve(
    input: string,
    available: readonly FeatureDefinition[],
    signal?: AbortSignal,
  ): Promise<IntentResult> {
    signal?.throwIfAborted();
    const normalized = input
      .normalize('NFKC')
      .toLowerCase()
      .replace(/[−–—]/g, '-')
      // Keep separated numeric tokens separate when removing ordinary word spacing.
      .replace(/(\d)\s+(?=[\d.])/g, '$1,')
      .replace(/\s+/g, '')
      .replace(/(?:[，,。.!?！？；;]|谢谢|就好|即可|吧|呀|啊)+$/g, '');
    // 多步请求必须逐项重新说，不能丢失参数后执行某个功能的默认值。
    const matches = available.filter((feature) =>
      feature.phrases.some((phrase) => normalized.includes(phrase)),
    );
    if (matches.length > 1)
      return { type: 'choices', featureIds: matches.map((feature) => feature.id) };
    const feature = matches[0];
    if (!feature) return { type: 'unknown' };
    const cancelledBold = feature.id === 'text.bold' && /取消(?:加粗|粗体)/.test(normalized);
    const request = cancelledBold ? normalized.replace(/取消(?:加粗|粗体)/g, '') : normalized;
    if (negatedRequest.test(request))
      return clarify('这句话包含取消或否定要求，请单独说清一项操作；目前支持“取消加粗”。');
    if (multipleSteps.test(normalized))
      return clarify('请一次只说一个操作和一个目标数值，例如“行距设为 1.5 倍”。');

    switch (feature.id) {
      case 'sheet.average':
        return averageIntent(normalized, feature.id);
      case 'text.enlarge':
        return fontSizeIntent(normalized, feature.id);
      case 'paragraph.spacing':
        return spacingIntent(normalized, feature.id);
      case 'paragraph.indent':
        return indentIntent(normalized, feature.id);
      case 'text.bold':
        if (!wholeRequest(normalized, '取消(?:加粗|粗体)|加粗|(?:(?:设置|设|改|变)(?:为|成))?粗体'))
          return clarify('加粗只有开启和取消两种状态，请说“加粗”或“取消加粗”。');
        return { type: 'match', featureId: feature.id, parameters: { bold: !cancelledBold } };
      case 'paragraph.center':
        if (!wholeRequest(normalized, '(?:(?:设置|设|改|调整|调)(?:为|成))?居中(?:对齐)?'))
          return clarify('段落居中不需要数值，请单独说“段落居中”。');
        return { type: 'match', featureId: feature.id };
      default:
        return { type: 'match', featureId: feature.id };
    }
  }
}
