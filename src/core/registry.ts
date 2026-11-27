import type { FeatureDefinition, HostInfo } from './contracts';
import { HelperError } from './errors';

export class FeatureRegistry {
  private readonly features = new Map<string, FeatureDefinition>();

  constructor(
    definitions: readonly FeatureDefinition[],
    private readonly flags: Readonly<Record<string, boolean>> = {},
  ) {
    for (const definition of definitions) {
      if (this.features.has(definition.id)) throw new Error(`重复的功能 ID：${definition.id}`);
      this.features.set(
        definition.id,
        Object.freeze({
          ...definition,
          kinds: Object.freeze([...definition.kinds]),
          requires: Object.freeze([...definition.requires]),
          phrases: Object.freeze([...definition.phrases]),
          readOptions: definition.readOptions
            ? Object.freeze({ ...definition.readOptions })
            : undefined,
        }),
      );
    }
  }

  list(): FeatureDefinition[] {
    return [...this.features.values()].filter((feature) => this.flags[feature.id] !== false);
  }

  get(id: string): FeatureDefinition {
    const feature = this.features.get(id);
    if (!feature || this.flags[id] === false)
      throw new HelperError('UNAVAILABLE', '这个功能目前没有开启。');
    return feature;
  }

  unavailableReason(feature: FeatureDefinition, host: HostInfo): string | undefined {
    if (!feature.kinds.includes(host.kind))
      return host.kind === 'text' ? '请在表格里使用' : '请在文字文档里使用';
    if (feature.requires.some((capability) => !host.capabilities.includes(capability)))
      return '当前版本尚未接通这个功能';
    return undefined;
  }
}
