import type { FeatureDefinition } from '../core/contracts';

export type IntentResult =
  | { type: 'match'; featureId: string }
  | { type: 'choices'; featureIds: string[] }
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
    if (matches.length === 1 && matches[0]) return { type: 'match', featureId: matches[0].id };
    if (matches.length > 1)
      return { type: 'choices', featureIds: matches.map((feature) => feature.id) };
    return { type: 'unknown' };
  }
}
