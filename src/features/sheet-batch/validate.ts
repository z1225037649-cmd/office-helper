import type { SheetBatchSource } from './types';

/** Validate the structural information needed to map values by exact header name. */
export function validateSources(sources: readonly SheetBatchSource[]): void {
  if (!Array.isArray(sources) || sources.length === 0) {
    throw new TypeError('Provide at least one sheet or file.');
  }

  const labels = new Set<string>();
  for (const source of sources) {
    if (!source || typeof source.sourceLabel !== 'string' || source.sourceLabel.trim() === '') {
      throw new TypeError('Each source must have a non-blank label.');
    }
    if (labels.has(source.sourceLabel)) {
      throw new TypeError(`Source labels must be unique: ${source.sourceLabel}`);
    }
    labels.add(source.sourceLabel);

    if (!Array.isArray(source.headers) || source.headers.length === 0) {
      throw new TypeError(`Source "${source.sourceLabel}" must have at least one header.`);
    }
    const headers = new Set<string>();
    for (const header of source.headers) {
      if (typeof header !== 'string' || header.trim() === '') {
        throw new TypeError(`Source "${source.sourceLabel}" has a blank or invalid header.`);
      }
      if (headers.has(header)) {
        throw new TypeError(`Source "${source.sourceLabel}" has a duplicate header: ${header}`);
      }
      headers.add(header);
    }

    if (!Array.isArray(source.rows)) {
      throw new TypeError(`Source "${source.sourceLabel}" must have a rows array.`);
    }
    for (let rowIndex = 0; rowIndex < source.rows.length; rowIndex += 1) {
      const row = source.rows[rowIndex];
      if (!Array.isArray(row)) {
        throw new TypeError(`Row ${rowIndex} in source "${source.sourceLabel}" is invalid.`);
      }
      if (row.length > source.headers.length) {
        throw new TypeError(
          `Row ${rowIndex} in source "${source.sourceLabel}" has more cells than headers.`,
        );
      }
    }
  }
}
