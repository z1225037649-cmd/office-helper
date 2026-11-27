import type { CellType, CellValue } from '../../core/contracts';
import type { WpsApplication, WpsSelection } from './bridge';

function ordinaryType(value: CellValue): CellType {
  if (value === null) return 'empty';
  if (typeof value === 'number') return Number.isFinite(value) ? 'number' : 'unknown';
  if (typeof value === 'string') return 'text';
  if (typeof value === 'boolean') return 'boolean';
  return 'unknown';
}

function sameValues(raw: WpsSelection['Value2'], expected: CellValue[][]): boolean {
  const rows = expected.length;
  const columns = expected[0]?.length ?? 0;
  const current = Array.isArray(raw) ? raw : rows === 1 && columns === 1 ? [[raw]] : undefined;
  return (
    !!current &&
    current.length === rows &&
    current.every(
      (row, r) =>
        Array.isArray(row) &&
        row.length === columns &&
        row.every((value, c) => Object.is(value, expected[r]?.[c])),
    )
  );
}

/**
 * Error variants can look like ordinary numbers after Value2 crosses the bridge.
 * Ask the native function about each native cell reference before classifying it.
 * Unsupported or incomplete probes omit details, so the average feature refuses them.
 */
export function readWpsSheetDetails(
  selection: WpsSelection,
  app: WpsApplication,
  values: CellValue[][],
): { cellTypes: CellType[][]; rowScope: 'all' } | Record<string, never> {
  try {
    const cells = selection.Cells;
    const functions = app.WorksheetFunction;
    if (
      selection.Areas?.Count !== 1 ||
      !cells ||
      typeof cells.Item !== 'function' ||
      !functions ||
      typeof functions.IsError !== 'function'
    )
      return {};

    const cellTypes: CellType[][] = [];
    for (let row = 0; row < values.length; row++) {
      const types: CellType[] = [];
      const source = values[row]!;
      for (let column = 0; column < source.length; column++) {
        const cell = cells.Item(row + 1, column + 1);
        if (!cell) return {};
        const error = functions.IsError(cell);
        if (![true, false, -1, 0].includes(error)) return {};
        if (!Object.is(cell.Value2, source[column])) return {};
        types.push(error === true || error === -1 ? 'error' : ordinaryType(source[column]!));
      }
      cellTypes.push(types);
    }

    // Value2 is read again to reject a recalculation or changed range during probes.
    if (!sameValues(selection.Value2, values)) return {};
    // A single rectangular Range is read in full; no visible-only filtering is applied.
    return { cellTypes, rowScope: 'all' };
  } catch {
    return {};
  }
}
