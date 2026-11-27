import { enlargeText } from './enlarge-text/feature';
import { sumSelection } from './sum-selection/feature';
import { averageSelection } from './average-selection/feature';
import { inspectSelection } from './inspect-selection/feature';
import type { FeatureDefinition } from '../core/contracts';
import { boldText, centerParagraph, spaceParagraph, indentParagraph } from './text-format/features';
import { autofitColumns, wrapCells, emphasizeHeader } from './sheet-layout/feature';
import { printCheck } from './print-check/feature';
import { saveCopy, exportPdf } from './save-copy/feature';
import { checkRecords } from './check-records/feature';
import { noticeTemplate } from './notice-template/feature';
import { normalizeDocumentStyles } from './document-styles/feature';
import { bidChecklist } from './bid-checklist/feature';

// 新功能在自己的目录实现，然后在这里注册；界面自动生成入口。
export const features: readonly FeatureDefinition[] = [
  enlargeText,
  boldText,
  centerParagraph,
  spaceParagraph,
  indentParagraph,
  sumSelection,
  averageSelection,
  inspectSelection,
  autofitColumns,
  wrapCells,
  emphasizeHeader,
  printCheck,
  saveCopy,
  exportPdf,
  checkRecords,
  noticeTemplate,
  normalizeDocumentStyles,
  bidChecklist,
];
