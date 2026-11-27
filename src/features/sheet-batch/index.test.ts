import { describe, expect, it } from 'vitest';
import { analyzeSheetNormalization, createSheetBatchPlan } from './index';
import type { SheetBatchSource } from './index';

describe('standalone sheet batch preview', () => {
  it('unions exact header names, aligns values, appends rows, and keeps source row provenance', () => {
    const sources: SheetBatchSource[] = [
      {
        sourceLabel: 'January',
        headers: ['Name', 'Amount'],
        rows: [
          ['Ada', 10],
          ['Lin', 20],
        ],
      },
      { sourceLabel: 'February', headers: ['Amount', 'Memo'], rows: [[30, 'paid'], [40]] },
    ];

    expect(createSheetBatchPlan(sources)).toEqual({
      columns: ['Name', 'Amount', 'Memo'],
      sources: [
        { sourceLabel: 'January', rowCount: 2, missingColumns: ['Memo'], extraColumns: [] },
        { sourceLabel: 'February', rowCount: 2, missingColumns: ['Name'], extraColumns: ['Memo'] },
      ],
      rowCountSummary: { sourceCount: 2, inputRowCount: 4, plannedRowCount: 4 },
      rows: [
        { values: ['Ada', 10, undefined], sourceLabel: 'January', originalRowIndex: 0 },
        { values: ['Lin', 20, undefined], sourceLabel: 'January', originalRowIndex: 1 },
        { values: [undefined, 30, 'paid'], sourceLabel: 'February', originalRowIndex: 0 },
        { values: [undefined, 40, undefined], sourceLabel: 'February', originalRowIndex: 1 },
      ],
    });
  });

  it('preserves order, value types, and input data without mutating any source', () => {
    const sources: SheetBatchSource[] = [
      { sourceLabel: 'First', headers: ['value'], rows: [[0], [false], [null], ['']] },
      { sourceLabel: 'Second', headers: ['value'], rows: [['  unchanged  ']] },
    ];
    const before = structuredClone(sources);

    const plan = createSheetBatchPlan(sources);

    expect(plan.rows.map((row) => row.values[0])).toEqual([0, false, null, '', '  unchanged  ']);
    expect(sources).toEqual(before);
  });

  it('rejects ambiguous headers, duplicate source labels, and cells without header names', () => {
    expect(() =>
      createSheetBatchPlan([{ sourceLabel: 'A', headers: ['Name', 'Name'], rows: [] }]),
    ).toThrow('duplicate header');
    expect(() =>
      createSheetBatchPlan([
        { sourceLabel: 'A', headers: ['Name'], rows: [] },
        { sourceLabel: 'A', headers: ['Name'], rows: [] },
      ]),
    ).toThrow('unique');
    expect(() =>
      createSheetBatchPlan([{ sourceLabel: 'A', headers: ['Name'], rows: [['Ada', 10]] }]),
    ).toThrow('more cells than headers');
  });
});

describe('read-only normalization analysis', () => {
  it('reports whitespace, varied blanks, invalid and mixed ISO dates, and mixed value types', () => {
    const sources: SheetBatchSource[] = [
      {
        sourceLabel: 'People',
        headers: ['Name', 'Date', 'Score'],
        rows: [
          [' Ada ', '2024-02-29', 12],
          ['   ', '2024-02-30', '12'],
          [null, 'not a date', 15],
          ['', '2024-03-01', 18],
        ],
      },
    ];
    const before = structuredClone(sources);

    const analysis = analyzeSheetNormalization(sources);

    expect(analysis.findingCount).toBe(analysis.findings.length);
    expect(analysis.findings).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          kind: 'surrounding-whitespace',
          columnName: 'Name',
          rowIndex: 0,
        }),
        expect.objectContaining({
          kind: 'blank-representation',
          columnName: 'Name',
          rowIndex: null,
          blankRepresentations: ['whitespace-string', 'null', 'empty-string'],
        }),
        expect.objectContaining({
          kind: 'invalid-date-like-value',
          columnName: 'Date',
          rowIndex: 1,
        }),
        expect.objectContaining({
          kind: 'mixed-date-like-values',
          columnName: 'Date',
          rowIndex: null,
        }),
        expect.objectContaining({
          kind: 'mixed-value-types',
          columnName: 'Score',
          observedTypes: ['number', 'text'],
        }),
      ]),
    );
    expect(sources).toEqual(before);
  });

  it('accepts consistent blank values and valid homogeneous ISO dates without findings', () => {
    const sources: SheetBatchSource[] = [
      {
        sourceLabel: 'Dates',
        headers: ['Date', 'Note'],
        rows: [
          ['2024-02-29', null],
          ['2024-03-01', null],
        ],
      },
    ];
    expect(analyzeSheetNormalization(sources)).toEqual({ findings: [], findingCount: 0 });
  });

  it('reports non-finite numbers but never proposes a replacement', () => {
    const result = analyzeSheetNormalization([
      { sourceLabel: 'Values', headers: ['Value'], rows: [[Infinity]] },
    ]);
    expect(result.findings).toEqual([
      expect.objectContaining({ kind: 'non-finite-number', rowIndex: 0 }),
    ]);
    expect(result.findings[0]).not.toHaveProperty('replacementValue');
  });

  it('bounds stored review findings and related row samples while keeping full counts', () => {
    const blanks = Array.from({ length: 600 }, (_, index) => [index % 2 ? '' : null]);
    const whitespace = Array.from({ length: 600 }, (_, index) => [` x${index} `]);
    const analysis = analyzeSheetNormalization([
      { sourceLabel: 'Blanks', headers: ['Name'], rows: blanks },
      { sourceLabel: 'Whitespace', headers: ['Name'], rows: whitespace },
    ]);
    expect(analysis.findings).toHaveLength(500);
    expect(analysis.findingCount).toBe(601);
    expect(analysis.findings[0]?.relatedRowIndexes).toHaveLength(200);
    expect(analysis.findings[0]?.relatedRowIndexCount).toBe(600);
  });
});
