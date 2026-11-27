import type {
  CellValue,
  DuplicateOptions,
  FeatureDefinition,
  ReportFinding,
  SelectionSnapshot,
} from '../../core/contracts';
import { HelperError } from '../../core/errors';

const findingLimit = 200;

function requireRows(selection: SelectionSnapshot): CellValue[][] {
  if (selection.kind !== 'sheet')
    throw new HelperError('UNAVAILABLE', '请在表格中选择要检查的记录。');
  const rows = selection.values;
  const columns = rows?.[0]?.length ?? 0;
  if (!Array.isArray(rows) || !rows.length || !columns)
    throw new HelperError('EMPTY_SELECTION', '请先选择包含记录的连续矩形区域。');
  if (
    Array.from(rows).some(
      (row) =>
        !Array.isArray(row) ||
        row.length !== columns ||
        Array.from(row).some(
          (value) =>
            value !== null &&
            typeof value !== 'string' &&
            typeof value !== 'boolean' &&
            !(typeof value === 'number' && Number.isFinite(value)),
        ),
    )
  ) {
    throw new HelperError(
      'INVALID_PLAN',
      '记录数据不完整或包含无法可靠比较的值，请重新选择矩形区域。',
    );
  }
  const types = selection.cellTypes;
  if (
    selection.rowScope !== 'all' ||
    !Array.isArray(types) ||
    types.length !== rows.length ||
    Array.from(types).some(
      (row, r) =>
        !Array.isArray(row) ||
        row.length !== columns ||
        Array.from(row).some((type, c) => {
          const value = rows[r]![c];
          return !(
            type === 'error' ||
            (type === 'empty' && (value === null || value === '')) ||
            (type === 'number' && typeof value === 'number') ||
            (type === 'text' && typeof value === 'string') ||
            (type === 'boolean' && typeof value === 'boolean')
          );
        }),
    )
  ) {
    throw new HelperError(
      'UNAVAILABLE',
      '无法可靠读取完整选区及每格类型，暂不能区分错误值与普通记录。请使用支持类型读取的版本后重新检查。',
    );
  }
  return rows;
}

function validateOptions(options: DuplicateOptions | undefined, columns: number): DuplicateOptions {
  if (
    !options ||
    !Array.isArray(options.keyColumns) ||
    !Array.isArray(options.requiredColumns) ||
    typeof options.trimWhitespace !== 'boolean' ||
    typeof options.ignoreCase !== 'boolean' ||
    typeof options.hasHeader !== 'boolean'
  ) {
    throw new HelperError(
      'INVALID_PLAN',
      '请选择判断列或必填列，并明确表头、首尾空格和大小写规则。',
    );
  }
  if (!options.keyColumns.length && !options.requiredColumns.length)
    throw new HelperError('INVALID_PLAN', '请至少选择一个判断列或必填列。');
  for (const selected of [options.keyColumns, options.requiredColumns]) {
    if (
      Array.from(selected).some(
        (column) => !Number.isSafeInteger(column) || column < 0 || column >= columns,
      ) ||
      new Set(selected).size !== selected.length
    ) {
      throw new HelperError(
        'INVALID_PLAN',
        '所选列必须是选区内有效且不重复的列（列索引从 0 开始）。',
      );
    }
  }
  return options;
}

function isBlank(value: CellValue): boolean {
  return value === null || (typeof value === 'string' && !value.trim());
}

function typedKey(value: CellValue, options: DuplicateOptions): [string, CellValue] {
  if (typeof value !== 'string') return [typeof value, value];
  let normalized = options.trimWhitespace ? value.trim() : value;
  if (options.ignoreCase) normalized = normalized.toLowerCase();
  return ['string', normalized];
}

function columnList(columns: number[]): string {
  return columns.length ? columns.map((column) => `第 ${column + 1} 列`).join('、') : '未启用';
}

export const checkRecords: FeatureDefinition = {
  id: 'sheet.checkRecords',
  title: '检查重复与漏填',
  description: '按所选判断列和必填列，只读检查疑似重复记录与空白必填项',
  icon: 'inspect',
  kinds: ['sheet'],
  requires: ['selection.read'],
  readOptions: { includeCellTypes: true },
  form: 'duplicates',
  phrases: [
    '检查重复与漏填',
    '检查重复',
    '检查漏填',
    '重复记录',
    '查重',
    '查漏填',
    '找出重复和漏填',
    '找重复和漏填',
    '找出重复',
    '找出漏填',
  ],
  prepare(selection, parameters) {
    const rows = requireRows(selection);
    const options = validateOptions(parameters?.duplicates, rows[0]!.length);
    const start = options.hasHeader ? 1 : 0;
    const firstRows = new Map<string, number>();
    const duplicateGroups = new Set<string>();
    const findings: ReportFinding[] = [];
    let duplicates = 0;
    let missingCells = 0;
    let missingRows = 0;
    let skippedKeys = 0;
    let errorRows = 0;
    let errorCells = 0;
    const addFinding = (finding: ReportFinding) => {
      if (findings.length < findingLimit) findings.push(finding);
    };

    for (let row = start; row < rows.length; row += 1) {
      const values = rows[row]!;
      const rowErrors = selection.cellTypes![row]!.filter((type) => type === 'error').length;
      if (rowErrors) {
        errorRows += 1;
        errorCells += rowErrors;
        continue;
      }
      if (options.keyColumns.length) {
        const keyValues = options.keyColumns.map((column) => values[column]!);
        if (keyValues.some(isBlank)) {
          skippedKeys += 1;
        } else {
          // JSON 元组保留各值类型和复合键边界，不用拼接分隔符或强制转字符串。
          const key = JSON.stringify(keyValues.map((value) => typedKey(value, options)));
          const first = firstRows.get(key);
          if (first !== undefined) {
            duplicates += 1;
            duplicateGroups.add(key);
            addFinding({
              id: `duplicate-${row}`,
              label: `疑似重复：选区第 ${row + 1} 行`,
              description: `判断列（${columnList(options.keyColumns)}）的值和类型按当前规则与选区第 ${first + 1} 行相同。首次出现行不计为重复；请人工核对，不会自动删除。`,
              row,
              column: options.keyColumns[0]!,
            });
          } else {
            firstRows.set(key, row);
          }
        }
      }
      let rowMissing = false;
      for (const column of options.requiredColumns) {
        if (!isBlank(values[column]!)) continue;
        missingCells += 1;
        rowMissing = true;
        addFinding({
          id: `missing-${row}-${column}`,
          label: `疑似漏填：选区第 ${row + 1} 行第 ${column + 1} 列`,
          description:
            '此列被设为必填，但该格为空值、空字符串或全空白字符。请核对是否确实需要填写，源数据未修改。',
          row,
          column,
        });
      }
      if (rowMissing) missingRows += 1;
    }

    const dataRows = rows.length - start;
    const totalFindings = duplicates + missingCells;
    const rules = `列号均相对当前选区：判断列 ${columnList(options.keyColumns)}；必填列 ${columnList(options.requiredColumns)}。${options.hasHeader ? '首行为表头，不参与检查' : '无表头，首行参与检查'}；字符串比较${options.trimWhitespace ? '去除' : '保留'}首尾空格、${options.ignoreCase ? '忽略' : '区分'}大小写，保留原始值类型（数字 1 与文字“1”不同）。复合键须所有判断列均相同；若同时选择姓名和联系方式，同名但联系方式不同不视为重复。空白指空值、空串或全空白字符，0 和 false 不算空白；任一判断列空白时跳过该行判重。含错误单元格的记录整行排除，不参与判重或漏填检查。`;
    const counts = `已扫描 ${dataRows} 行数据；其中 ${errorRows} 行含 ${errorCells} 个错误单元格，整行排除，实际检查 ${dataRows - errorRows} 行。疑似重复 ${duplicates} 行（各组首次出现之后的行，共 ${duplicateGroups.size} 组）；疑似漏填 ${missingCells} 格（涉及 ${missingRows} 行）。${options.keyColumns.length ? `因判断列空白跳过判重 ${skippedKeys} 行。` : ''}`;
    const display =
      totalFindings > findingLimit
        ? `共 ${totalFindings} 条结果，仅显示前 ${findingLimit} 条；计数已覆盖全部 ${dataRows} 行数据。`
        : `共 ${totalFindings} 条结果，已全部显示。`;
    return {
      summary: `只读检查「${selection.targetLabel}」。${rules}${counts}不修改、排序或删除源数据。`,
      plan: {
        type: 'report',
        message: `${counts}${display}${rules}结果位置按原始选区计算；检查包含所选范围中的隐藏或筛选行。不修改、排序或删除源数据。`,
        findings,
      },
    };
  },
};
