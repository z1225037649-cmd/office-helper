import { enlargeText } from './enlarge-text/feature';
import { sumSelection } from './sum-selection/feature';
import { inspectSelection } from './inspect-selection/feature';
import type { FeatureDefinition } from '../core/contracts';

// 新功能在自己的目录实现，然后在这里注册；界面自动生成入口。
export const features: readonly FeatureDefinition[] = [enlargeText, sumSelection, inspectSelection];
