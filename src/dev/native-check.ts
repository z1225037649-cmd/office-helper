import { createAdapter } from '../adapters';
import { createRuntime } from '../app/runtime';
import { userMessage } from '../core/errors';
import { getWpsApplication } from '../adapters/wps/bridge';
import { readWpsParagraphs } from '../adapters/wps/paragraphs';
import type { HostAdapter } from '../core/contracts';
import { checkWpsFormats } from './wps-format-check';
import { checkWordFormats } from './word-format-check';
import { checkSheetAverage } from './sheet-average-check';

const paragraphFixture = 'OFFICE_HELPER_NATIVE_TEST\r第二段：居中验收\r第三段：保持原样';

// 仅本机开发验证入口。只允许本工程生成的、含固定标记的独立测试文件。
const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';
const hostName = params.get('host') ?? '';
const status = document.getElementById('status')!;
async function report(stage: string, details: Record<string, unknown> = {}) {
  status.textContent = `${stage}\n${JSON.stringify(details, null, 2)}`;
  const response = await fetch('/__native-report', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'X-Helper-Test': token },
    body: JSON.stringify({ host: hostName, stage, ...details }),
  });
  if (!response.ok) throw new Error('本地测试入口未启用，停止执行。');
}

async function checkWpsParagraphBoundaries(adapter: HostAdapter) {
  const app = getWpsApplication('text');
  const selected = app.Selection as typeof app.Selection & {
    SetRange(start: number, end: number): void;
  };
  const original = await adapter.readSelection();
  if (original.text?.trim() !== paragraphFixture || !selected?.Range || !original.paragraphs)
    throw new Error('段落边界验收只允许专用三段样本');
  const full = selected.Range.Duplicate;
  const before = JSON.stringify(original.paragraphs);
  const firstEnd = original.paragraphs[0]!.text.length;
  const secondEnd = firstEnd + original.paragraphs[1]!.text.length;
  const all = () => readWpsParagraphs(full, app)?.snapshots;
  const expectOriginal = () => {
    if (JSON.stringify(all()) !== before) throw new Error('非目标段落或原始格式发生变化');
  };
  const { runner } = createRuntime(adapter);
  try {
    selected.SetRange(2, 6);
    let task = await runner.prepare('paragraph.center');
    if (
      !task.summary.includes('1 个完整段落') ||
      !task.summary.includes('OFFICE_HELPER_NATIVE_TEST'.slice(0, 24))
    )
      throw new Error('部分选字预览没有显示完整段落');
    runner.cancel(task.id);
    expectOriginal();
    await report('center-cancel-verified');
    task = await runner.prepare('paragraph.center');
    await runner.execute(task.id);
    const partial = all();
    if (partial?.map((p) => p.alignment).join() !== 'Centered,Right,Left')
      throw new Error('部分选字居中越过目标段落');
    selected.SetRange(secondEnd, full.End);
    await adapter.readSelection();
    await runner.undo();
    expectOriginal();
    await report('center-partial-moved-undo-verified', { after: partial.map((p) => p.alignment) });

    selected.SetRange(2, secondEnd);
    task = await runner.prepare('paragraph.center');
    if (!task.summary.includes('2 个完整段落')) throw new Error('独占结束边界包含了下一段');
    await runner.execute(task.id);
    if (
      all()
        ?.map((p) => p.alignment)
        .join() !== 'Centered,Centered,Left'
    )
      throw new Error('结束边界后的段落被修改');
    await runner.undo();
    expectOriginal();
    await report('center-exclusive-end-verified');

    selected.SetRange(2, 6);
    task = await runner.prepare('paragraph.center');
    const format = selected.Range.Paragraphs!.Item(1).Range.ParagraphFormat!;
    const initialAlignment = format.Alignment;
    const right =
      app.Enum?.wdAlignParagraphRight ?? app.Enum?.WdParagraphAlignment?.wdAlignParagraphRight;
    if (!Number.isInteger(right)) throw new Error('右对齐枚举不可读');
    try {
      format.Alignment = right;
      let rejected = false;
      try {
        await runner.execute(task.id);
      } catch (error) {
        rejected = (error as { code?: string }).code === 'STALE_SELECTION';
      }
      if (!rejected || format.Alignment !== right) throw new Error('旧预览没有被安全拦截');
      await report('center-stale-preview-verified');
    } finally {
      format.Alignment = initialAlignment;
    }
    expectOriginal();

    task = await runner.prepare('paragraph.center');
    await runner.execute(task.id);
    try {
      format.Alignment = right;
      let rejected = false;
      try {
        await runner.undo();
      } catch (error) {
        rejected = (error as { code?: string }).code === 'STALE_SELECTION';
      }
      if (!rejected || format.Alignment !== right) throw new Error('恢复覆盖了后续格式');
      await report('center-restore-conflict-verified');
    } finally {
      format.Alignment = initialAlignment;
    }
    expectOriginal();

    selected.SetRange(0, 0);
    if ((await adapter.readSelection()).text !== '') throw new Error('插入光标被当成文字选区');
    for (const id of ['text.enlarge', 'text.bold', 'paragraph.center']) {
      let rejected = false;
      try {
        await runner.prepare(id);
      } catch (error) {
        rejected = (error as { code?: string }).code === 'EMPTY_SELECTION';
      }
      if (!rejected) throw new Error('插入光标操作未被拒绝');
    }
    expectOriginal();
    await report('collapsed-selection-verified');
  } finally {
    selected.SetRange(full.Start, full.End);
  }
}

async function run() {
  if (!import.meta.env.DEV || !token) throw new Error('仅用于本机开发验证');
  await report('page-loaded', { userAgent: navigator.userAgent });
  const adapter = await createAdapter(location.search);
  await report('host-connected', {
    kind: adapter.info.kind,
    capabilities: adapter.info.capabilities,
  });
  try {
    let selection = await adapter.readSelection();
    const expectedName =
      adapter.info.kind === 'text'
        ? 'OfficeHelper-Native-Test.docx'
        : 'OfficeHelper-Native-Test.xlsx';
    for (let attempt = 0; attempt < 500; attempt++) {
      const validName = selection.documentName === expectedName;
      const marker =
        adapter.info.kind === 'text'
          ? ['OFFICE_HELPER_NATIVE_TEST', paragraphFixture].includes(selection.text?.trim() ?? '')
          : JSON.stringify(selection.values) === '[[120],[85],[45.5],[200]]';
      if (validName && marker) break;
      if (attempt === 0)
        await report('waiting-for-selection', { kind: selection.kind, nameMatch: validName });
      if (attempt === 499)
        throw new Error(
          `测试文件或选区标记不匹配，未执行修改。nameMatch=${validName}, markerMatch=${marker}, selectedLength=${selection.text?.length ?? 0}`,
        );
      await new Promise((resolve) => setTimeout(resolve, 600));
      selection = await adapter.readSelection();
    }
    await report('selection-verified', {
      kind: selection.kind,
      fontSize: selection.fontSize,
      rows: selection.values?.length,
    });
    await report('claim', {
      run: `${hostName}:${selection.kind}:${params.get('run') ?? params.get('check') ?? 'default'}`,
    });
    if (params.get('check') === 'sheet-average') {
      await checkSheetAverage(adapter, report);
      return;
    }
    if (hostName === 'wps' && selection.kind === 'text') {
      const app = getWpsApplication('text');
      const range = app.Selection?.Range;
      const safe = (read: () => unknown) => {
        try {
          return read();
        } catch {
          return 'unsupported';
        }
      };
      await report('paragraph-probe', {
        enumSources: safe(() => {
          const globals = window as unknown as Record<string, Record<string, unknown> | undefined>;
          const sources = {
            application: app.Enum,
            wps: globals.wps?.Enum,
            global: globals.Enum,
            window: window,
          } as Record<string, unknown>;
          const keys = [
            'wdAlignParagraphLeft',
            'wdAlignParagraphCenter',
            'wdAlignParagraphRight',
            'wdAlignParagraphJustify',
            'wdMainTextStory',
            'wdListNoNumbering',
          ];
          return Object.fromEntries(
            Object.entries(sources).map(([label, value]) => [
              label,
              {
                type: typeof value,
                values: Object.fromEntries(
                  keys.map((key) => [
                    key,
                    safe(() => (value as Record<string, unknown> | undefined)?.[key]),
                  ]),
                ),
              },
            ]),
          );
        }),
        alignmentEnums: safe(() =>
          Object.fromEntries(
            [
              'wdAlignParagraphLeft',
              'wdAlignParagraphCenter',
              'wdAlignParagraphRight',
              'wdAlignParagraphJustify',
            ].map((key) => [
              key,
              (app.Enum?.WdParagraphAlignment ?? app.Enum)?.[key as 'wdAlignParagraphLeft'],
            ]),
          ),
        ),
        mainStory: safe(() => app.Enum?.WdStoryType?.wdMainTextStory ?? app.Enum?.wdMainTextStory),
        noList: safe(() => app.Enum?.WdListType?.wdListNoNumbering ?? app.Enum?.wdListNoNumbering),
        story: safe(() => range?.StoryType),
        tables: safe(() => range?.Tables?.Count),
        count: safe(() => range?.Paragraphs?.Count),
        first: safe(() => {
          const p = range?.Paragraphs?.Item(1).Range;
          return {
            start: p?.Start,
            end: p?.End,
            alignment: p?.ParagraphFormat?.Alignment,
            spacing: p?.ParagraphFormat?.LineSpacing,
            spacingRule: p?.ParagraphFormat?.LineSpacingRule,
            firstIndent: p?.ParagraphFormat?.FirstLineIndent,
            leftIndent: p?.ParagraphFormat?.LeftIndent,
            rightIndent: p?.ParagraphFormat?.RightIndent,
            list: p?.ListFormat?.ListType,
          };
        }),
        snapshots: selection.paragraphs,
      });
    }
    const { runner } = createRuntime(adapter);
    if (selection.kind === 'text' && params.get('check') === 'paragraph-format') {
      if (hostName === 'wps') await checkWpsFormats(adapter, report);
      else if (hostName === 'office') await checkWordFormats(adapter, report);
      else throw new Error('格式专项验收必须从真实Word或WPS任务窗格运行');
      return;
    }
    if (selection.kind === 'text') {
      const originalSize = selection.fontSize;
      const task = await runner.prepare('text.enlarge');
      await runner.execute(task.id);
      const after = await adapter.readSelection();
      await report('resize-verified', { before: originalSize, after: after.fontSize });
      if (!after.fontSize || after.fontSize === originalSize) throw new Error('字号修改没有生效');
      await runner.undo();
      const restored = await adapter.readSelection();
      if (restored.fontSize !== originalSize || restored.text !== selection.text)
        throw new Error('恢复校验失败');
      if (hostName === 'wps' && adapter.info.capabilities.includes('text.bold')) {
        const originalBold = restored.bold;
        if (typeof originalBold !== 'boolean') throw new Error('测试粗细读取失败');
        const checkText = async (bold: boolean) => {
          const state = await adapter.readSelection();
          if (
            state.text !== selection.text ||
            state.fontSize !== originalSize ||
            state.bold !== bold
          )
            throw new Error('加粗测试的正文、字号或粗细不匹配');
          return state;
        };
        const applyBold = async (bold: boolean) => {
          const task = await runner.prepare('text.bold', { bold });
          await runner.execute(task.id);
          await checkText(bold);
        };
        // Verify both directions and each direction's undo, restoring the fixture at the end.
        await applyBold(!originalBold);
        await runner.undo();
        await checkText(originalBold);
        await report('bold-undo-verified', {
          before: originalBold,
          after: !originalBold,
          restored: originalBold,
        });
        await applyBold(!originalBold);
        await applyBold(originalBold);
        await runner.undo();
        await checkText(!originalBold);
        await applyBold(originalBold);
        await report('unbold-undo-verified', {
          before: !originalBold,
          after: originalBold,
          restored: !originalBold,
          final: originalBold,
        });
      }
      if (hostName === 'wps' && adapter.info.capabilities.includes('paragraph.center')) {
        const before = await adapter.readSelection();
        const task = await runner.prepare('paragraph.center');
        if (!task.summary.includes('完整段落')) throw new Error('居中预览未说明完整段落范围');
        await runner.execute(task.id);
        const centered = await adapter.readSelection();
        if (
          !centered.paragraphs?.length ||
          centered.paragraphs.some((p) => p.alignment !== 'Centered')
        )
          throw new Error('段落居中没有生效');
        await runner.undo();
        const restored = await adapter.readSelection();
        if (
          JSON.stringify(restored.paragraphs) !== JSON.stringify(before.paragraphs) ||
          restored.text !== before.text ||
          restored.fontSize !== before.fontSize ||
          restored.bold !== before.bold
        )
          throw new Error('段落逐段恢复失败');
        await report('center-undo-verified', {
          count: before.paragraphs?.length,
          before: before.paragraphs?.map((p) => p.alignment),
          after: centered.paragraphs.map((p) => p.alignment),
          restored: restored.paragraphs?.map((p) => p.alignment),
        });
        if (selection.text?.trim() === paragraphFixture) await checkWpsParagraphBoundaries(adapter);
      }
      await report('passed', {
        kind: 'text',
        ...(hostName === 'office' ? { coverage: 'font-size-only' } : {}),
        before: originalSize,
        after: after.fontSize,
        restored: restored.fontSize,
        ...(hostName === 'wps' && adapter.info.capabilities.includes('text.bold')
          ? { bold: 'both-directions-and-undo' }
          : {}),
        ...(hostName === 'wps' && adapter.info.capabilities.includes('paragraph.center')
          ? { center: 'full-paragraphs-and-individual-undo' }
          : {}),
      });
    } else {
      const task = await runner.prepare('sheet.sum');
      const result = await runner.execute(task.id);
      if (!result.message.includes('450.5')) throw new Error('合计不正确');
      await report('passed', { kind: 'sheet', total: 450.5, rows: selection.values?.length });
    }
  } finally {
    adapter.dispose();
  }
}

void run().catch(async (error) => {
  try {
    await report('failed', {
      message: userMessage(error),
      detail: error instanceof Error ? error.message : String(error),
      code: error?.code ?? '',
    });
  } catch {
    status.textContent = '测试入口已停止。请打开正常的小帮手入口。';
  }
});
