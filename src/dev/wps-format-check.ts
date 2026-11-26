import { getWpsApplication } from '../adapters/wps/bridge';
import { readWpsParagraphs } from '../adapters/wps/paragraphs';
import { createRuntime } from '../app/runtime';
import type { HostAdapter } from '../core/contracts';

/** Token-gated caller only; never run against a user's document. */
export async function checkWpsFormats(
  adapter: HostAdapter,
  report: (stage: string, details?: Record<string, unknown>) => Promise<void>,
) {
  const app = getWpsApplication('text');
  const selection = app.Selection as typeof app.Selection & {
    SetRange(start: number, end: number): void;
  };
  const before = await adapter.readSelection();
  if (
    before.documentName !== 'OfficeHelper-Native-Test.docx' ||
    before.text?.trim() !== 'OFFICE_HELPER_NATIVE_TEST\r第二段：居中验收\r第三段：保持原样' ||
    before.paragraphs?.length !== 3 ||
    !selection?.Range
  )
    throw new Error('格式验收只允许专用三段样本');
  const full = selection.Range.Duplicate;
  const original = readWpsParagraphs(full, app)!;
  const originalRevision = original.revision;
  const e = app.Enum?.WdLineSpacing ?? app.Enum!;
  const first = original.targets[0]!.range.ParagraphFormat!;
  const second = original.targets[1]!.range.ParagraphFormat!;
  const third = original.targets[2]!.range.ParagraphFormat!;
  const { runner } = createRuntime(adapter);
  const state = () => readWpsParagraphs(full, app)!;
  const assertRevision = (revision: string) => {
    if (state().revision !== revision) throw new Error('段落原值或非目标格式不匹配');
  };
  try {
    // Cover mixed native rules and both character/point indent originals.
    first.CharacterUnitFirstLineIndent = 2;
    second.FirstLineIndent = 12;
    third.FirstLineIndent = -12;
    second.LineSpacingRule = e.wdLineSpaceMultiple;
    second.LineSpacing = 15;
    third.LineSpacingRule = e.wdLineSpaceDouble;
    const mixed = state();
    await report('format-mixed-originals', { paragraphs: mixed.snapshots });
    for (const multiple of [1, 1.5, 2, 1.2, 3]) {
      selection.SetRange(full.Start, full.End);
      let task = await runner.prepare('paragraph.spacing', { lineSpacingMultiple: multiple });
      runner.cancel(task.id);
      assertRevision(mixed.revision);
      task = await runner.prepare('paragraph.spacing', { lineSpacingMultiple: multiple });
      await runner.execute(task.id);
      const after = state();
      if (after.snapshots.some((p) => Math.abs(p.lineSpacing - multiple * 12) > 0.001))
        throw new Error('行距写入不符');
      selection.SetRange(2, 6);
      await runner.undo();
      assertRevision(mixed.revision);
      await report('spacing-individual-undo-verified', {
        multiple,
        rules: after.targets.map((p) => p.spacingRule),
      });
    }
    for (const points of [0, 24, 28.35, 144]) {
      selection.SetRange(full.Start, full.End);
      let task = await runner.prepare('paragraph.indent', { firstLineIndentPoints: points });
      runner.cancel(task.id);
      assertRevision(mixed.revision);
      task = await runner.prepare('paragraph.indent', { firstLineIndentPoints: points });
      await runner.execute(task.id);
      if (state().snapshots.some((p) => p.firstLineIndent !== points))
        throw new Error('缩进写入不符');
      selection.SetRange(2, 6);
      await runner.undo();
      assertRevision(mixed.revision);
      await report('indent-character-and-point-undo-verified', { points });
    }
    // Character-unit indentation must keep its units when the font size changes.
    selection.SetRange(full.Start, full.End);
    const sizeTask = await runner.prepare('text.enlarge');
    await runner.execute(sizeTask.id);
    await runner.undo();
    assertRevision(mixed.revision);
    await report('resize-character-indent-regression-verified');
    for (const [id, parameters] of [
      ['paragraph.spacing', { lineSpacingMultiple: 1.2 }],
      ['paragraph.indent', { firstLineIndentPoints: 28.35 }],
    ] as const) {
      const secondEnd = original.targets[1]!.end;
      for (const end of [6, secondEnd]) {
        selection.SetRange(2, end);
        const task = await runner.prepare(id, parameters);
        await runner.execute(task.id);
        const changed = state();
        const count = end === 6 ? 1 : 2;
        if (
          changed.targets
            .slice(count)
            .some(
              (p, i) => JSON.stringify(p.snapshot) !== JSON.stringify(mixed.snapshots[i + count]),
            )
        )
          throw new Error('修改越过完整段落或独占结束边界');
        selection.SetRange(secondEnd, full.End);
        await runner.undo();
        assertRevision(mixed.revision);
      }
      selection.SetRange(2, 6);
      let task = await runner.prepare(id, parameters);
      const old = first.Alignment;
      const right =
        app.Enum?.WdParagraphAlignment?.wdAlignParagraphRight ?? app.Enum?.wdAlignParagraphRight;
      try {
        first.Alignment = right;
        let rejected = false;
        try {
          await runner.execute(task.id);
        } catch (error) {
          rejected = (error as { code?: string }).code === 'STALE_SELECTION';
        }
        if (!rejected || first.Alignment !== right) throw new Error('旧预览未被拒绝');
      } finally {
        first.Alignment = old;
      }
      assertRevision(mixed.revision);
      task = await runner.prepare(id, parameters);
      await runner.execute(task.id);
      try {
        first.Alignment = right;
        let rejected = false;
        try {
          await runner.undo();
        } catch (error) {
          rejected = (error as { code?: string }).code === 'STALE_SELECTION';
        }
        if (!rejected || first.Alignment !== right) throw new Error('恢复覆盖了后续编辑');
      } finally {
        first.Alignment = old;
        first.CharacterUnitFirstLineIndent = 0;
        first.FirstLineIndent = mixed.targets[0]!.rawIndent;
        first.CharacterUnitFirstLineIndent = mixed.targets[0]!.characterIndent;
        first.LineSpacingRule = mixed.targets[0]!.spacingRule;
      }
      assertRevision(mixed.revision);
      await report('format-boundaries-conflicts-verified', { id });
    }
  } finally {
    for (const target of original.targets) {
      const f = target.range.ParagraphFormat!;
      f.CharacterUnitFirstLineIndent = 0;
      f.FirstLineIndent = target.rawIndent;
      if (target.characterIndent !== 0) f.CharacterUnitFirstLineIndent = target.characterIndent;
      f.LineSpacingRule = target.spacingRule;
      if (target.spacingRule === original.multipleRule) f.LineSpacing = target.rawSpacing;
    }
    selection.SetRange(full.Start, full.End);
    assertRevision(originalRevision);
  }
  await report('spacing-indent-passed', { restored: true });
}
