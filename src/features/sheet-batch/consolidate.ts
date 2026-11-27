import type { SheetBatchPlan, SheetBatchSource } from './types';
import { validateSources } from './validate';

/**
 * Build a read-only union/append preview using exact, case-sensitive header names.
 * The first source defines the reference schema for each source's `extraColumns`.
 */
export function createSheetBatchPlan(sources: readonly SheetBatchSource[]): SheetBatchPlan {
  validateSources(sources);

  const columns: string[] = [];
  const columnSet = new Set<string>();
  for (const source of sources) {
    for (const header of source.headers) {
      if (!columnSet.has(header)) {
        columnSet.add(header);
        columns.push(header);
      }
    }
  }

  const referenceHeaders = new Set(sources[0]!.headers);
  const summaries = sources.map((source) => {
    const sourceHeaders = new Set(source.headers);
    return {
      sourceLabel: source.sourceLabel,
      rowCount: source.rows.length,
      missingColumns: columns.filter((column) => !sourceHeaders.has(column)),
      extraColumns: source.headers.filter((column) => !referenceHeaders.has(column)),
    };
  });

  const rows = sources.flatMap((source) => {
    const sourceIndexes = new Map(source.headers.map((header, index) => [header, index]));
    return source.rows.map((row, originalRowIndex) => ({
      values: columns.map((column) => {
        const sourceIndex = sourceIndexes.get(column);
        return sourceIndex === undefined ? undefined : row[sourceIndex];
      }),
      sourceLabel: source.sourceLabel,
      originalRowIndex,
    }));
  });

  return {
    columns,
    sources: summaries,
    rowCountSummary: {
      sourceCount: sources.length,
      inputRowCount: rows.length,
      plannedRowCount: rows.length,
    },
    rows,
  };
}
