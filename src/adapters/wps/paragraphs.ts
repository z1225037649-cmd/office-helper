import type { ParagraphSnapshot, TextActionPlan } from '../../core/contracts';
import { MAX_SELECTION_PARAGRAPHS, MAX_SELECTION_TEXT } from '../../core/selection';
import type { WpsApplication, WpsRange } from './bridge';

interface ParagraphTarget {
  range: WpsRange;
  start: number;
  end: number;
  alignment: number;
  spacingRule: number;
  rawSpacing: number;
  rawIndent: number;
  characterIndent?: number;
  snapshot: ParagraphSnapshot;
}

export interface WpsParagraphState {
  targets: ParagraphTarget[];
  snapshots: ParagraphSnapshot[];
  revision: string;
  center: number;
  multipleRule?: number;
  presetRules?: number[];
}

function ruleForMultiple(state: WpsParagraphState, multiple: number): number {
  const index = [1, 1.5, 2].indexOf(multiple);
  return index < 0 ? state.multipleRule! : state.presetRules![index]!;
}

// Desktop bridges can expose single-precision points. Keep twip values stable without
// hiding meaningful changes (the smallest supported input is 0.05 pt).
function points(value: number): number {
  const twips = Math.round(value * 20) / 20;
  return Math.abs(value - twips) < 0.0005 ? twips : value;
}

export function paragraphRevision(targets: ParagraphTarget[]): string {
  return JSON.stringify(
    targets.map((target) => [
      target.start,
      target.end,
      target.alignment,
      target.spacingRule,
      points(target.rawSpacing),
      target.characterIndent,
      target.snapshot,
    ]),
  );
}

export function projectedParagraphRevision(state: WpsParagraphState, plan: TextActionPlan): string {
  return paragraphRevision(
    state.targets.map((target) => {
      const next = { ...target, snapshot: { ...target.snapshot } };
      if (plan.type === 'paragraph.center') {
        next.alignment = state.center;
        next.snapshot.alignment = 'Centered';
      } else if (plan.type === 'paragraph.spacing') {
        next.spacingRule = ruleForMultiple(state, plan.multiple);
        next.rawSpacing = next.snapshot.lineSpacing = points(plan.multiple * 12);
      } else if (plan.type === 'paragraph.indent') {
        next.snapshot.firstLineIndent = plan.points;
        next.characterIndent = 0;
      }
      return next;
    }),
  );
}

export function writeWpsParagraphs(state: WpsParagraphState, plan: TextActionPlan): void {
  for (const target of state.targets) {
    const format = target.range.ParagraphFormat!;
    if (plan.type === 'paragraph.center') format.Alignment = state.center;
    else if (plan.type === 'paragraph.spacing') {
      format.LineSpacingRule = ruleForMultiple(state, plan.multiple);
      if (ruleForMultiple(state, plan.multiple) === state.multipleRule)
        format.LineSpacing = plan.multiple * 12;
    } else if (plan.type === 'paragraph.indent') {
      // WPS retains a nonzero character indent when only the points are assigned.
      format.CharacterUnitFirstLineIndent = 0;
      format.FirstLineIndent = plan.points;
    }
  }
}

export function restoreWpsParagraphs(state: WpsParagraphState, type: TextActionPlan['type']): void {
  for (const target of state.targets) {
    const format = target.range.ParagraphFormat!;
    if (type === 'paragraph.center') format.Alignment = target.alignment;
    else if (type === 'paragraph.spacing') {
      format.LineSpacingRule = target.spacingRule;
      // Preset rules restore their own value; assigning LineSpacing can convert a
      // preset into Multiple on some bridges. Only Multiple needs a second write.
      if (target.spacingRule === state.multipleRule) format.LineSpacing = target.rawSpacing;
    } else if (type === 'paragraph.indent') {
      format.CharacterUnitFirstLineIndent = 0;
      format.FirstLineIndent = target.rawIndent;
      if (target.characterIndent !== 0)
        format.CharacterUnitFirstLineIndent = target.characterIndent;
    }
  }
}

/** Synchronous desktop JSAPI only. Missing interfaces leave existing character tools usable. */
export function readWpsParagraphs(
  range: WpsRange | undefined,
  app: WpsApplication,
): WpsParagraphState | undefined {
  try {
    // Current WPS desktop exposes flat Enum keys; some bridges group them by enum type.
    const alignments = app.Enum?.WdParagraphAlignment ?? app.Enum;
    const mainStory = app.Enum?.WdStoryType?.wdMainTextStory ?? app.Enum?.wdMainTextStory;
    const noList = app.Enum?.WdListType?.wdListNoNumbering ?? app.Enum?.wdListNoNumbering;
    const spacing = app.Enum?.WdLineSpacing ?? app.Enum;
    const spacingRules = [
      spacing?.wdLineSpaceSingle,
      spacing?.wdLineSpace1pt5,
      spacing?.wdLineSpaceDouble,
      spacing?.wdLineSpaceMultiple,
    ];
    const knownSpacingRules =
      spacingRules.every(Number.isInteger) && new Set(spacingRules).size === 4;
    if (!range || !alignments || !Number.isInteger(mainStory) || !Number.isInteger(noList))
      return undefined;
    const names = [
      ['wdAlignParagraphLeft', 'Left'],
      ['wdAlignParagraphCenter', 'Centered'],
      ['wdAlignParagraphRight', 'Right'],
      ['wdAlignParagraphJustify', 'Justified'],
    ] as const;
    const values = names.map(([key]) => alignments[key]);
    if (values.some((value) => !Number.isInteger(value)) || new Set(values).size !== 4)
      return undefined;
    if (range.StoryType !== mainStory || range.Tables?.Count !== 0 || range.Start >= range.End)
      return undefined;
    const paragraphs = range.Paragraphs;
    if (!paragraphs || !Number.isInteger(paragraphs.Count) || paragraphs.Count < 1)
      return undefined;
    // Some clients include the paragraph at the exclusive End boundary.
    if (paragraphs.Count > MAX_SELECTION_PARAGRAPHS + 1) return undefined;
    const targets: ParagraphTarget[] = [];
    let length = 0;
    for (let index = 1; index <= paragraphs.Count; index++) {
      const full = paragraphs.Item(index).Range.Duplicate;
      if (!Number.isInteger(full.Start) || !Number.isInteger(full.End) || full.Start >= full.End)
        return undefined;
      if (full.Start >= range.End || full.End <= range.Start) continue;
      if (full.StoryType !== mainStory || full.Tables?.Count !== 0 || full.Text.includes('\u0007'))
        return undefined;
      if (
        (full.Document.FullName || full.Document.Name) !==
        (range.Document.FullName || range.Document.Name)
      )
        return undefined;
      if (targets.length && targets[targets.length - 1]!.end !== full.Start) return undefined;
      const format = full.ParagraphFormat;
      const alignment = format?.Alignment;
      const name = names.find(([key]) => alignments[key] === alignment)?.[1];
      if (
        !format ||
        !name ||
        full.ListFormat?.ListType !== noList ||
        ![format.LineSpacing, format.FirstLineIndent, format.LeftIndent, format.RightIndent].every(
          Number.isFinite,
        ) ||
        !Number.isInteger(format.LineSpacingRule)
      )
        return undefined;
      length += full.Text.length;
      if (targets.length >= MAX_SELECTION_PARAGRAPHS || length > MAX_SELECTION_TEXT)
        return undefined;
      const presetIndex = knownSpacingRules ? spacingRules.indexOf(format.LineSpacingRule) : -1;
      const normalizedSpacing =
        presetIndex >= 0 && presetIndex < 3
          ? [12, 18, 24][presetIndex]!
          : points(format.LineSpacing!);
      let characterIndent: number | undefined;
      try {
        const value = format.CharacterUnitFirstLineIndent;
        if (Number.isFinite(value)) characterIndent = value;
      } catch {
        /* Older bridges keep other paragraph tools available. */
      }
      targets.push({
        range: full,
        start: full.Start,
        end: full.End,
        alignment: alignment!,
        spacingRule: format.LineSpacingRule!,
        rawSpacing: format.LineSpacing!,
        rawIndent: format.FirstLineIndent!,
        characterIndent,
        snapshot: {
          text: full.Text,
          alignment: name,
          lineSpacing: normalizedSpacing,
          spacingMode:
            presetIndex >= 0 &&
            normalizedSpacing > 0 &&
            (presetIndex === 3 || points(format.LineSpacing!) === normalizedSpacing) &&
            // WPS canonicalizes these Multiple values to preset rules on write.
            !(presetIndex === 3 && [12, 18, 24].includes(normalizedSpacing))
              ? 'multiple'
              : 'unsupported',
          firstLineIndent: points(format.FirstLineIndent!),
          firstLineIndentRestorable:
            characterIndent !== undefined &&
            Math.abs(characterIndent) <= 1000 &&
            Math.abs(format.FirstLineIndent!) <= 1584,
          leftIndent: points(format.LeftIndent!),
          rightIndent: points(format.RightIndent!),
          isListItem: false,
        },
      });
    }
    if (
      !targets.length ||
      targets[0]!.start > range.Start ||
      targets[targets.length - 1]!.end < range.End
    )
      return undefined;
    return {
      targets,
      snapshots: targets.map((target) => target.snapshot),
      revision: paragraphRevision(targets),
      center: alignments.wdAlignParagraphCenter!,
      multipleRule: knownSpacingRules ? spacing!.wdLineSpaceMultiple : undefined,
      presetRules: knownSpacingRules ? (spacingRules.slice(0, 3) as number[]) : undefined,
    };
  } catch {
    return undefined;
  }
}
