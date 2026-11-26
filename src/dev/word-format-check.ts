import { explicitMultipleSpacing } from '../adapters/office/spacing';
import { createRuntime } from '../app/runtime';
import type { FeatureParameters, HostAdapter, PreparedTask } from '../core/contracts';

const fixtureText = 'OFFICE_HELPER_NATIVE_TEST\r第二段：居中验收\r第三段：保持原样';
const fixtureName = 'OfficeHelper-Native-Test.docx';

interface NativeParagraph {
  text: string;
  alignment: string;
  lineSpacing: number;
  spacingMode: 'multiple' | 'unsupported';
  firstLineIndent: number;
  leftIndent: number;
  rightIndent: number;
  isListItem: boolean;
  fontSize: number;
  bold: boolean;
}

interface NativeState {
  text: string;
  paragraphs: NativeParagraph[];
}

interface FormatCase {
  id: string;
  parameters?: FeatureParameters;
  property: 'alignment' | 'lineSpacing' | 'firstLineIndent';
  value: string | number;
}

const center: FormatCase = {
  id: 'paragraph.center',
  property: 'alignment',
  value: 'Centered',
};
const spacing = (multiple: number): FormatCase => ({
  id: 'paragraph.spacing',
  parameters: { lineSpacingMultiple: multiple },
  property: 'lineSpacing',
  value: multiple * 12,
});
const indent = (points: number): FormatCase => ({
  id: 'paragraph.indent',
  parameters: { firstLineIndentPoints: points },
  property: 'firstLineIndent',
  value: points,
});

function assertState(actual: NativeState, expected: NativeState, label: string) {
  if (actual.text !== expected.text || actual.paragraphs.length !== expected.paragraphs.length)
    throw new Error(`${label}：正文或段落数量发生变化`);
  actual.paragraphs.forEach((paragraph, index) => {
    const original = expected.paragraphs[index]!;
    for (const key of Object.keys(original) as (keyof NativeParagraph)[]) {
      const a = paragraph[key];
      const b = original[key];
      if (typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) < 0.001) continue;
      if (a !== b) throw new Error(`${label}：第${index + 1}段的${key}不符（${a} / ${b}）`);
    }
  });
}

function project(original: NativeState, test: FormatCase, count = 3): NativeState {
  const expected = structuredClone(original);
  for (const paragraph of expected.paragraphs.slice(0, count))
    Object.assign(paragraph, { [test.property]: test.value });
  return expected;
}

/** Called only by the development, token-gated entry point, on a generated fixture. */
export async function checkWordFormats(
  adapter: HostAdapter,
  report: (stage: string, details?: Record<string, unknown>) => Promise<void>,
) {
  if (
    !import.meta.env.DEV ||
    adapter.info.id !== 'office' ||
    adapter.info.kind !== 'text' ||
    typeof Word === 'undefined' ||
    Office.context.host !== Office.HostType.Word ||
    !Office.context.requirements.isSetSupported('WordApi', '1.3')
  )
    throw new Error('Word格式验收需要真实Word任务窗格和WordApi 1.3');
  const selected = await adapter.readSelection();
  if (
    selected.documentName !== fixtureName ||
    selected.text?.trim() !== fixtureText ||
    selected.paragraphs?.length !== 3
  )
    throw new Error('Word格式验收只允许全选专用三段测试样本');
  const documentUrl = Office.context.document.url;
  if (!documentUrl) throw new Error('Word测试文件必须先保存，才能确认文档归属');
  let guardedText: string | undefined;

  const guardIdentity = () => {
    if (Office.context.document.url !== documentUrl)
      throw new Error('Word测试文档已切换，停止验收和清理');
  };
  const inFixture = async <T>(
    operation: (
      context: Word.RequestContext,
      body: Word.Body,
      paragraphs: Word.Paragraph[],
    ) => Promise<T>,
  ): Promise<T> => {
    guardIdentity();
    return Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      body.load('text');
      paragraphs.load('items');
      await context.sync();
      guardIdentity();
      if (
        body.text.trim() !== fixtureText ||
        (guardedText !== undefined && body.text !== guardedText) ||
        paragraphs.items.length !== 3
      )
        throw new Error('专用样本正文标记或段落发生变化，停止验收和清理');
      return operation(context, body, paragraphs.items);
    });
  };
  const readNative = () =>
    inFixture(async (context, body, paragraphs): Promise<NativeState> => {
      const ranges = paragraphs.map((paragraph) => paragraph.getRange('Whole'));
      for (const [index, paragraph] of paragraphs.entries()) {
        paragraph.load([
          'text',
          'alignment',
          'lineSpacing',
          'firstLineIndent',
          'leftIndent',
          'rightIndent',
          'isListItem',
        ]);
        ranges[index]!.font.load(['size', 'bold']);
      }
      const xml = body.getOoxml();
      await context.sync();
      const multiple = explicitMultipleSpacing(
        xml.value,
        paragraphs.map((paragraph) => paragraph.lineSpacing),
      );
      return {
        text: body.text,
        paragraphs: paragraphs.map((paragraph, index) => ({
          text: paragraph.text,
          alignment: paragraph.alignment,
          lineSpacing: paragraph.lineSpacing,
          spacingMode: multiple[index] ? 'multiple' : 'unsupported',
          firstLineIndent: paragraph.firstLineIndent,
          leftIndent: paragraph.leftIndent,
          rightIndent: paragraph.rightIndent,
          isListItem: paragraph.isListItem,
          fontSize: ranges[index]!.font.size,
          bold: ranges[index]!.font.bold,
        })),
      };
    });
  const select = (part: 'all' | 'partial' | 'exclusive-end' | 'third' | 'collapsed') =>
    inFixture(async (context, body, paragraphs) => {
      if (part === 'all') body.getRange('Whole').select();
      else if (part === 'third') paragraphs[2]!.getRange('Content').select();
      else if (part === 'collapsed') paragraphs[0]!.getRange('Start').select();
      else {
        const matches = paragraphs[0]!.search('FICE', { matchCase: true, matchWholeWord: false });
        matches.load('items');
        await context.sync();
        if (matches.items.length !== 1) throw new Error('部分选字的唯一样本标记未找到');
        const range = matches.items[0]!;
        if (part === 'partial') range.select();
        else range.getRange('Start').expandTo(paragraphs[2]!.getRange('Start')).select();
      }
      await context.sync();
    });
  const setFirstAlignment = (alignment: string) =>
    inFixture(async (context, _body, paragraphs) => {
      paragraphs[0]!.alignment = alignment as Word.Alignment;
      await context.sync();
    });
  const original = await readNative();
  guardedText = original.text;
  const expectedOriginals = [
    { alignment: 'Left', lineSpacing: 12, firstLineIndent: 0, leftIndent: 6, rightIndent: 8 },
    { alignment: 'Right', lineSpacing: 18, firstLineIndent: 12, leftIndent: 0, rightIndent: 4 },
    { alignment: 'Left', lineSpacing: 24, firstLineIndent: -12, leftIndent: 12, rightIndent: 0 },
  ];
  for (const [index, paragraph] of original.paragraphs.entries()) {
    if (
      paragraph.fontSize !== 18 ||
      paragraph.bold !== false ||
      paragraph.isListItem ||
      paragraph.spacingMode !== 'multiple' ||
      Object.entries(expectedOriginals[index]!).some(
        ([key, value]) => paragraph[key as keyof NativeParagraph] !== value,
      )
    )
      throw new Error('请重新生成Word专项样本：需要三段不同原格式、显式倍数行距和18磅常规文字');
  }
  await report('word-format-fixture-verified', { paragraphs: original.paragraphs });
  const { runner } = createRuntime(adapter);
  const verify = async (expected: NativeState, label: string) =>
    assertState(await readNative(), expected, label);
  const prepare = async (id: string, parameters?: FeatureParameters) => {
    await readNative();
    return runner.prepare(id, parameters);
  };
  const execute = async (task: PreparedTask) => {
    await readNative();
    return runner.execute(task.id);
  };
  const undo = async () => {
    await readNative();
    await runner.undo();
  };
  const reject = async (operation: () => Promise<unknown>, code: string) => {
    try {
      await operation();
    } catch (error) {
      if ((error as { code?: string }).code === code) return;
      throw error;
    }
    throw new Error(`验收要求拒绝${code}，实际却执行成功`);
  };
  // Cleanup writes only attributes deliberately exercised by this diagnostic.
  // It refuses to touch a different document or an edited fixture body.
  const restoreOriginal = () =>
    inFixture(async (context, _body, paragraphs) => {
      for (const [index, paragraph] of paragraphs.entries()) {
        const before = original.paragraphs[index]!;
        const range = paragraph.getRange('Whole');
        range.font.size = before.fontSize;
        range.font.bold = before.bold;
        paragraph.alignment = before.alignment as Word.Alignment;
        paragraph.lineSpacing = before.lineSpacing;
        paragraph.firstLineIndent = before.firstLineIndent;
      }
      await context.sync();
    });
  let failure: unknown;
  try {
    await select('all');
    let task = await prepare('text.enlarge', { fontSize: 22 });
    runner.cancel(task.id);
    await reject(() => execute(task), 'EXPIRED_TASK');
    await verify(original, '字号取消');
    task = await prepare('text.enlarge', { fontSize: 22 });
    await execute(task);
    const enlarged = structuredClone(original);
    enlarged.paragraphs.forEach((paragraph) => (paragraph.fontSize = 22));
    await verify(enlarged, '字号修改');
    await undo();
    await verify(original, '字号恢复');
    await report('word-resize-undo-verified');

    const bold = structuredClone(original);
    bold.paragraphs.forEach((paragraph) => (paragraph.bold = true));
    task = await prepare('text.bold', { bold: true });
    runner.cancel(task.id);
    await reject(() => execute(task), 'EXPIRED_TASK');
    await verify(original, '加粗取消');
    await execute(await prepare('text.bold', { bold: true }));
    await verify(bold, '加粗');
    await undo();
    await verify(original, '加粗恢复');
    await execute(await prepare('text.bold', { bold: true }));
    task = await prepare('text.bold', { bold: false });
    runner.cancel(task.id);
    await reject(() => execute(task), 'EXPIRED_TASK');
    await verify(bold, '取消加粗的预览取消');
    await execute(await prepare('text.bold', { bold: false }));
    await verify(original, '取消加粗');
    await undo();
    await verify(bold, '取消加粗恢复');
    await execute(await prepare('text.bold', { bold: false }));
    await verify(original, '粗细最终原值');
    await report('word-bold-both-directions-undo-verified');

    for (const test of [
      center,
      ...[1, 1.5, 2, 1.2, 3].map(spacing),
      ...[0, 24, 28.35, 144].map(indent),
    ]) {
      await select('all');
      task = await prepare(test.id, test.parameters);
      if (!task.summary.includes('3 个完整段落')) throw new Error('三段预览范围不符');
      runner.cancel(task.id);
      await reject(() => execute(task), 'EXPIRED_TASK');
      await verify(original, '完整段落取消');
      await execute(await prepare(test.id, test.parameters));
      await verify(project(original, test), '完整段落修改及非目标属性');
      await select('third');
      await adapter.readSelection();
      await undo();
      await verify(original, '移动选区后逐段恢复');
      await report('word-full-paragraph-individual-undo-verified', {
        feature: test.id,
        value: test.value,
      });
    }

    for (const test of [center, spacing(1.2), indent(28.35)]) {
      for (const part of ['partial', 'exclusive-end'] as const) {
        const count = part === 'partial' ? 1 : 2;
        await select(part);
        task = await prepare(test.id, test.parameters);
        if (
          !task.summary.includes(`${count} 个完整段落`) ||
          !task.summary.includes('OFFICE_HELPER_NATIVE_TEST'.slice(0, 24))
        )
          throw new Error('部分选字或独占结束边界的完整段落预览不符');
        runner.cancel(task.id);
        await reject(() => execute(task), 'EXPIRED_TASK');
        await verify(original, '边界预览取消');
        await execute(await prepare(test.id, test.parameters));
        await verify(project(original, test, count), '边界范围及未选中段落');
        await select('third');
        await adapter.readSelection();
        await undo();
        await verify(original, '边界修改移动选区恢复');
        await report('word-paragraph-boundary-verified', { feature: test.id, part, count });
      }

      await select('partial');
      task = await prepare(test.id, test.parameters);
      await select('third');
      await reject(() => execute(task), 'STALE_SELECTION');
      await verify(original, '选区变化拒绝');

      await select('partial');
      task = await prepare(test.id, test.parameters);
      const externallyAligned = structuredClone(original);
      externallyAligned.paragraphs[0]!.alignment = 'Right';
      try {
        await setFirstAlignment('Right');
        await reject(() => execute(task), 'STALE_SELECTION');
        await verify(externallyAligned, '过期预览不能覆盖新格式');
      } finally {
        await setFirstAlignment(original.paragraphs[0]!.alignment);
      }
      await verify(original, '过期预览场景恢复');

      task = await prepare(test.id, test.parameters);
      await execute(task);
      const conflicting = project(original, test, 1);
      conflicting.paragraphs[0]!.alignment = 'Right';
      try {
        await setFirstAlignment('Right');
        await reject(undo, 'STALE_SELECTION');
        await verify(conflicting, '恢复不能覆盖后续编辑');
      } finally {
        await restoreOriginal();
      }
      await verify(original, '恢复冲突场景清理');
      await report('word-preview-and-restore-conflicts-verified', { feature: test.id });
    }

    await select('collapsed');
    if ((await adapter.readSelection()).text !== '') throw new Error('Word插入光标被当成文字选区');
    for (const test of [
      { id: 'text.enlarge', parameters: { fontSize: 22 } },
      { id: 'text.bold', parameters: { bold: true } },
      center,
      spacing(1.5),
      indent(24),
    ])
      await reject(() => prepare(test.id, test.parameters), 'EMPTY_SELECTION');
    await verify(original, '插入光标拒绝');
    await report('word-collapsed-selection-verified');
  } catch (error) {
    failure = error;
  } finally {
    try {
      await restoreOriginal();
      await select('all');
      await verify(original, 'Word最终清理');
      await report('word-format-cleanup-verified', { restored: true });
    } catch (error) {
      await report('word-format-cleanup-failed', {
        restored: false,
        detail: error instanceof Error ? error.message : String(error),
      });
      failure ??= error;
    }
  }
  if (failure) throw failure;
  await report('word-format-passed', {
    restored: true,
    coverage: [
      'resize',
      'bold-and-unbold',
      'center',
      'explicit-multiple-spacing',
      'point-indent',
      'preview-cancel',
      'individual-originals',
      'partial-selection',
      'exclusive-end',
      'moved-selection-undo',
      'stale-selection-and-format',
      'restore-conflict',
      'collapsed-selection',
    ],
  });
}
