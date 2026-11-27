/** Primitive values accepted from a table reader. A missing cell is represented as undefined. */
export type SheetCellValue = string | number | boolean | null | undefined;

/** A named, header-based table. Row indexes in reports are zero-based within `rows`. */
export interface SheetBatchSource {
  readonly sourceLabel: string;
  readonly headers: readonly string[];
  readonly rows: readonly (readonly SheetCellValue[])[];
}

export interface SheetBatchSourceSummary {
  readonly sourceLabel: string;
  readonly rowCount: number;
  /** Columns from the union schema that this source does not provide. */
  readonly missingColumns: readonly string[];
  /** Columns this source has that are not present in the first source's header. */
  readonly extraColumns: readonly string[];
}

export interface SheetBatchPreviewRow {
  /** Values aligned with the plan's `columns`; absent source cells are undefined. */
  readonly values: readonly SheetCellValue[];
  readonly sourceLabel: string;
  /** Zero-based index in the source's `rows` array; the header is not counted. */
  readonly originalRowIndex: number;
}

export interface SheetBatchPlan {
  /** Exact header names, in first-seen order across the sources. */
  readonly columns: readonly string[];
  readonly sources: readonly SheetBatchSourceSummary[];
  readonly rowCountSummary: {
    readonly sourceCount: number;
    readonly inputRowCount: number;
    readonly plannedRowCount: number;
  };
  /** Appended in source order, preserving row order within each source. */
  readonly rows: readonly SheetBatchPreviewRow[];
}

export type NormalizationFindingKind =
  | 'surrounding-whitespace'
  | 'blank-representation'
  | 'invalid-date-like-value'
  | 'mixed-date-like-values'
  | 'mixed-value-types'
  | 'non-finite-number';

export type BlankRepresentation = 'empty-string' | 'whitespace-string' | 'null' | 'undefined';
export type ObservedValueType = 'text' | 'number' | 'boolean' | 'other';

/** A review-only observation. Suggested actions never contain replacement values. */
export interface NormalizationFinding {
  readonly kind: NormalizationFindingKind;
  readonly sourceLabel: string;
  readonly columnName: string;
  /** Zero-based data row index; null means the finding describes the column as a whole. */
  readonly rowIndex: number | null;
  readonly relatedRowIndexes?: readonly number[];
  /** Full number of related rows when the display sample is capped. */
  readonly relatedRowIndexCount?: number;
  readonly blankRepresentations?: readonly BlankRepresentation[];
  readonly observedTypes?: readonly ObservedValueType[];
  readonly suggestion: string;
}

export interface NormalizationAnalysis {
  readonly findings: readonly NormalizationFinding[];
  /** Total findings, including those omitted from the bounded preview array. */
  readonly findingCount: number;
}
