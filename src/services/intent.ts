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

export class LocalIntentProvider implements IntentProvider {
  async resolve(
    input: string,
    available: readonly FeatureDefinition[],
    signal?: AbortSignal,
  ): Promise<IntentResult> {
    signal?.throwIfAborted();
    const normalized = input.replace(/[\s，。！？、,.!?]/g, '').toLowerCase();
    // 多步请求交给用户逐项选择，避免只执行第一项。
    const matches = available.filter((feature) =>
      feature.phrases.some((phrase) => normalized.includes(phrase)),
    );
    if (matches.length === 1 && matches[0]) {
      if (matches[0].id !== 'text.enlarge') return { type: 'match', featureId: matches[0].id };

      const sizeMatch = normalized.match(
        /(?:放大(?:到|为|成)?|调整到|调到|改成|改为|设置为|设为|设到|字号(?:调整到|调到|改为|设为)?|到)(\d+(?:\.\d+)?)(磅|点|pt|厘米|毫米|英寸|cm|mm|in|px|%)?/i,
      );
      if (sizeMatch?.[1]) {
        const fontSize = Number(sizeMatch[1]);
        const unit = sizeMatch[2]?.toLowerCase();
        if (unit && !['磅', '点', 'pt'].includes(unit)) {
          return {
            type: 'clarify',
            message: '这里需要字号单位，请使用“磅”，例如“把字放大到 30 磅”。',
          };
        }
        if (Number.isFinite(fontSize) && fontSize >= 8 && fontSize <= 72) {
          return { type: 'match', featureId: matches[0].id, parameters: { fontSize } };
        }
        return { type: 'clarify', message: '字号请填写 8 到 72 磅之间的数字。' };
      }

      // Do not silently replace an explicit but incomplete or unsupported size request with the default increase.
      if (
        /(?:放大到|调整到|调到|改成|改为|设置为|设为|设到|字号|到)/.test(normalized) ||
        /(?:\d+(?:\.\d+)?(?:磅|点|号)|[一二三四五六七八九十百零]+(?:磅|号))/.test(normalized)
      ) {
        return { type: 'clarify', message: '请告诉我具体字号，例如“把字放大到 30 磅”。' };
      }
      return { type: 'match', featureId: matches[0].id };
    }
    if (matches.length > 1)
      return { type: 'choices', featureIds: matches.map((feature) => feature.id) };
    return { type: 'unknown' };
  }
}
