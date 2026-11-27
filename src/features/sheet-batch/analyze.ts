import type {
  BlankRepresentation,
  NormalizationAnalysis,
  NormalizationFinding,
  ObservedValueType,
  SheetBatchSource,
  SheetCellValue,
} from './types';
import { validateSources } from './validate';

const WHITESPACE_SUGGESTION = 'Review surrounding whitespace; the original value is unchanged.';
const BLANK_SUGGESTION = 'Review whether these blank representations should be consistent.';
const DATE_SUGGESTION = 'Verify the intended date manually; the original value is unchanged.';
const MIXED_SUGGESTION = 'Review the intended column format; no value conversion was applied.';
const maximumStoredFindings = 500;
const maximumRelatedRows = 200;

function sampledRows(groups: Iterable<readonly number[]>): {
  relatedRowIndexes: number[];
  relatedRowIndexCount: number;
} {
  const relatedRowIndexes: number[] = [];
  let relatedRowIndexCount = 0;
  for (const group of groups) {
    for (const rowIndex of group) {
      relatedRowIndexCount += 1;
      if (relatedRowIndexes.length < maximumRelatedRows) relatedRowIndexes.push(rowIndex);
    }
  }
  return { relatedRowIndexes, relatedRowIndexCount };
}

function blankRepresentation(value: SheetCellValue): BlankRepresentation | undefined {
  if (value === undefined) return 'undefined';
  if (value === null) return 'null';
  if (typeof value === 'string') {
    if (value === '') return 'empty-string';
    if (value.trim() === '') return 'whitespace-string';
  }
  return undefined;
}

function valueType(value: SheetCellValue): ObservedValueType {
  if (typeof value === 'string') return 'text';
  if (typeof value === 'number') return 'number';
  if (typeof value === 'boolean') return 'boolean';
  return 'other';
}

function isValidIsoDate(value: string): boolean {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/u.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (year < 1 || month < 1 || month > 12) return false;
  const leapYear = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth = [31, leapYear ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return day >= 1 && day <= daysInMonth[month - 1]!;
}

function hasIsoDateShape(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/u.test(value);
}

/**
 * Inspect string whitespace, inconsistent blank representations, strict ISO date values,
 * and mixed or non-finite values. This function reports review findings only and never
 * returns replacement values or changes its inputs.
 */
export function analyzeSheetNormalization(
  sources: readonly SheetBatchSource[],
): NormalizationAnalysis {
  validateSources(sources);
  const findings: NormalizationFinding[] = [];
  let findingCount = 0;
  const addFinding = (finding: NormalizationFinding) => {
    findingCount += 1;
    if (findings.length < maximumStoredFindings) findings.push(finding);
  };

  for (const source of sources) {
    for (let columnIndex = 0; columnIndex < source.headers.length; columnIndex += 1) {
      const columnName = source.headers[columnIndex]!;
      const blanks = new Map<BlankRepresentation, number[]>();
      const valueTypes = new Map<ObservedValueType, number[]>();
      const dateRows: number[] = [];
      const otherNonblankRows: number[] = [];

      for (let rowIndex = 0; rowIndex < source.rows.length; rowIndex += 1) {
        const value = source.rows[rowIndex]![columnIndex];
        const blankKind = blankRepresentation(value);
        if (blankKind) {
          const indexes = blanks.get(blankKind) ?? [];
          indexes.push(rowIndex);
          blanks.set(blankKind, indexes);
          if (typeof value === 'string' && value.trim() !== value) {
            addFinding({
              kind: 'surrounding-whitespace',
              sourceLabel: source.sourceLabel,
              columnName,
              rowIndex,
              suggestion: WHITESPACE_SUGGESTION,
            });
          }
          continue;
        }

        const type = valueType(value!);
        const typeIndexes = valueTypes.get(type) ?? [];
        typeIndexes.push(rowIndex);
        valueTypes.set(type, typeIndexes);

        if (typeof value === 'number' && !Number.isFinite(value)) {
          addFinding({
            kind: 'non-finite-number',
            sourceLabel: source.sourceLabel,
            columnName,
            rowIndex,
            suggestion: 'Review the numeric value; it is not finite and remains unchanged.',
          });
        }
        if (typeof value === 'string') {
          if (value.trim() !== value) {
            addFinding({
              kind: 'surrounding-whitespace',
              sourceLabel: source.sourceLabel,
              columnName,
              rowIndex,
              suggestion: WHITESPACE_SUGGESTION,
            });
          }
          if (hasIsoDateShape(value)) {
            if (!isValidIsoDate(value)) {
              addFinding({
                kind: 'invalid-date-like-value',
                sourceLabel: source.sourceLabel,
                columnName,
                rowIndex,
                suggestion: DATE_SUGGESTION,
              });
            } else {
              dateRows.push(rowIndex);
            }
          } else {
            otherNonblankRows.push(rowIndex);
          }
        } else {
          otherNonblankRows.push(rowIndex);
        }
      }

      if (blanks.size > 1) {
        const related = sampledRows(blanks.values());
        addFinding({
          kind: 'blank-representation',
          sourceLabel: source.sourceLabel,
          columnName,
          rowIndex: null,
          ...related,
          blankRepresentations: [...blanks.keys()],
          suggestion: BLANK_SUGGESTION,
        });
      }

      if (valueTypes.size > 1) {
        const related = sampledRows(valueTypes.values());
        addFinding({
          kind: 'mixed-value-types',
          sourceLabel: source.sourceLabel,
          columnName,
          rowIndex: null,
          ...related,
          observedTypes: [...valueTypes.keys()],
          suggestion: MIXED_SUGGESTION,
        });
      }

      if (dateRows.length > 0 && otherNonblankRows.length > 0) {
        const related = sampledRows([dateRows, otherNonblankRows]);
        addFinding({
          kind: 'mixed-date-like-values',
          sourceLabel: source.sourceLabel,
          columnName,
          rowIndex: null,
          ...related,
          suggestion: DATE_SUGGESTION,
        });
      }
    }
  }

  return { findings, findingCount };
}
