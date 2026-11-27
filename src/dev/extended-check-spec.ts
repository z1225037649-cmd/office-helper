import type { DocumentKind, FeatureParameters, SelectionSnapshot } from '../core/contracts';

export const extendedMarker = 'OFFICE_HELPER_EXTENDED_TEST_V1';
export const extendedMatrices = ['office-text', 'office-sheet', 'wps-text', 'wps-sheet'] as const;
export type ExtendedMatrix = (typeof extendedMatrices)[number];
export const extendedFeatures = [
  'sheet.autofit',
  'sheet.wrap',
  'sheet.header',
  'sheet.print',
  'document.saveCopy',
  'document.exportPdf',
  'sheet.checkRecords',
  'document.notice',
] as const;
export type ExtendedFeature = (typeof extendedFeatures)[number];

export interface ExtendedConfig {
  schema: 1;
  run: string;
  matrix: ExtendedMatrix;
  host: 'office' | 'wps';
  kind: DocumentKind;
  fixturePath: string;
  resultDirectory: string;
  fileName: string;
  text: string;
  values: (string | number | boolean)[][];
  notice: NonNullable<FeatureParameters['notice']>;
}

export const extendedRows = [
  ['编号', '姓名', '联系方式', '备注', '金额'],
  ['A-01', 'Alice', '1001', '常规', 10],
  ['A-02', ' alice ', '1001', '重复记录', 20],
  ['A-03', 'Bob', '', '漏填联系方式', 30],
  ['A-04', 'Carol', '1004', '', 0],
  ['A-05', 'Delta', '1005', '公式值：应保留公式而非只保留计算值', 30],
  ['A-06', 'ALICE', '1001', '隐藏行仍参与检查', 40],
  ['A-07', '', '1007', '漏填姓名，false不是空白', false],
];

export const normalizeText = (value: string) => value.replace(/\r\n?/gu, '\n').replace(/\n$/u, '');
export const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
export const addressOnly = (label: string) =>
  label.split('!').at(-1)!.replaceAll('$', '').toUpperCase();

export function fixtureText(run: string, matrix: string): string {
  return `${extendedMarker}\nRUN=${run}\nMATRIX=${matrix}\n仅供扩展能力隔离验收；原稿保持不变。`;
}

function winPath(path: string): string {
  const result = path.replaceAll('/', '\\');
  if (
    !/^[A-Za-z]:\\/u.test(result) ||
    /[\u0000-\u001f]/u.test(result) ||
    result.split('\\').some((part) => !part || part === '.' || part === '..')
  )
    throw new Error('只允许明确的本地绝对路径。');
  return result.toLowerCase();
}

export function assertExtendedConfig(config: ExtendedConfig): void {
  if (
    config.schema !== 1 ||
    !/^[a-f0-9]{32}$/u.test(config.run) ||
    !extendedMatrices.includes(config.matrix) ||
    config.matrix !== `${config.host}-${config.kind}` ||
    config.fileName !== `OfficeHelper-Extended-Test.${config.kind === 'text' ? 'docx' : 'xlsx'}` ||
    config.text !== fixtureText(config.run, config.matrix) ||
    !same(config.values, extendedRows)
  )
    throw new Error('扩展诊断配置、内容标记或宿主矩阵不匹配。');
  const suffix = `\\output\\native\\extended\\${config.run}`;
  const sourceSuffix = `${suffix}\\fixtures\\${config.matrix}\\${config.fileName.toLowerCase()}`;
  const source = winPath(config.fixturePath);
  if (!source.endsWith(sourceSuffix)) throw new Error('样本必须属于本轮独立目录。');
  const prefix = source.slice(0, -sourceSuffix.length);
  if (winPath(config.resultDirectory) !== `${prefix}${suffix}\\results\\${config.matrix}`)
    throw new Error('导出必须进入本轮专用新结果目录。');
  if (
    !config.notice ||
    config.notice.title !== `扩展诊断通知 ${config.run}` ||
    config.notice.time !== '2026年11月28日 09:00（测试数据）' ||
    config.notice.location !== '隔离验收室（测试数据）' ||
    config.notice.body !== `${extendedMarker}\nRUN=${config.run}\nMATRIX=${config.matrix}` ||
    config.notice.contact !== '测试联系人，不发送通知'
  )
    throw new Error('通知只能使用明确的隔离测试字段。');
}

export function assertSourcePath(actual: string, config: ExtendedConfig): void {
  let path = actual;
  if (/^file:/iu.test(path)) {
    const url = new URL(path);
    if (url.hostname && url.hostname !== 'localhost') throw new Error('不接受网络样本。');
    path = decodeURIComponent(url.pathname).replace(/^\/([A-Za-z]:)/u, '$1');
  }
  if (winPath(path) !== winPath(config.fixturePath))
    throw new Error('原生宿主报告的路径不是本轮专用样本。');
}

export function assertExtendedSelection(
  selection: SelectionSnapshot,
  config: ExtendedConfig,
  requireRange = true,
): void {
  assertExtendedConfig(config);
  if (
    !selection.documentId ||
    !selection.targetId ||
    selection.documentName !== config.fileName ||
    selection.kind !== config.kind
  )
    throw new Error('专用文件名、类型或身份不匹配。');
  if (!requireRange) return;
  if (
    config.kind === 'text'
      ? normalizeText(selection.text ?? '') !== config.text
      : addressOnly(selection.targetLabel) !== 'B3:F10' || !same(selection.values, extendedRows)
  )
    throw new Error('请选择全部四段测试正文，或 ExtendedData!B3:F10 精确矩阵。');
}

export function featureParameters(
  feature: ExtendedFeature,
  config: ExtendedConfig,
): FeatureParameters | undefined {
  if (feature === 'sheet.checkRecords')
    return {
      duplicates: {
        keyColumns: [1, 2],
        requiredColumns: [1, 2, 4],
        trimWhitespace: true,
        ignoreCase: true,
        hasHeader: true,
      },
    };
  if (feature === 'document.notice') return { notice: config.notice };
  if (feature === 'document.saveCopy' || feature === 'document.exportPdf')
    return {
      folder: config.resultDirectory,
      fileName: outputName(feature, config),
    };
  return undefined;
}

export function outputName(
  feature: 'document.saveCopy' | 'document.exportPdf',
  config: ExtendedConfig,
): string {
  return `Extended-${config.run}-${config.matrix}.${feature === 'document.exportPdf' ? 'pdf' : config.kind === 'text' ? 'docx' : 'xlsx'}`;
}

export function featuresFor(kind: DocumentKind): ExtendedFeature[] {
  return extendedFeatures.filter((id) =>
    kind === 'sheet' ? id !== 'document.notice' : id.startsWith('document.'),
  );
}

export const expectedFindings = [
  { id: 'duplicate-2', row: 2, column: 1, address: 'C5' },
  { id: 'missing-3-2', row: 3, column: 2, address: 'D6' },
  { id: 'duplicate-6', row: 6, column: 1, address: 'C9' },
  { id: 'missing-7-1', row: 7, column: 1, address: 'C10' },
];
