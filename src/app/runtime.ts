import type { HostAdapter } from '../core/contracts';
import { FeatureRegistry } from '../core/registry';
import { TaskRunner } from '../core/runner';
import { features } from '../features';
import { LocalIntentProvider, type IntentProvider } from '../services/intent';
import { appConfig } from './config';

/** 集中装配依赖；更换宿主或指令服务时不改界面和功能。 */
export function createRuntime(
  host: HostAdapter,
  intent: IntentProvider = new LocalIntentProvider(),
) {
  const registry = new FeatureRegistry(features, appConfig.featureFlags);
  return { host, registry, runner: new TaskRunner(host, registry), intent };
}

export type AppRuntime = ReturnType<typeof createRuntime>;
