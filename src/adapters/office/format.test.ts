import { afterEach, describe, expect, it, vi } from 'vitest';
import { OfficeAdapter } from './adapter';
import { FeatureRegistry } from '../../core/registry';
import { TaskRunner } from '../../core/runner';
import { features } from '../../features';

function setup() {
  const document = { url: 'file:///通知.docx' };
  const state = {
    text: '选中的几个字',
    size: 18,
    bold: false as boolean | null,
    selected: 'original',
    protected: false,
    writes: [] as string[],
    paragraphs: [
      {
        text: '第一段的完整内容',
        alignment: 'Left',
        lineSpacing: 12,
        firstLineIndent: 0,
        leftIndent: 6,
        rightIndent: 8,
        isListItem: false,
        rule: 'auto',
      },
      {
        text: '第二段的完整内容',
        alignment: 'Right',
        lineSpacing: 18,
        firstLineIndent: 12,
        leftIndent: 0,
        rightIndent: 4,
        isListItem: false,
        rule: 'auto',
      },
    ],
  };
  const check = (property: string) => {
    if (state.protected) throw new Error('Protected document');
    state.writes.push(property);
  };
  function rangeFor(id: string): unknown {
    const paragraphs = () =>
      state.paragraphs.map(
        (p) =>
          new Proxy(
            { ...p, load: vi.fn() },
            {
              get: (target, key) =>
                key in p ? p[key as keyof typeof p] : target[key as keyof typeof target],
              set: (_target, key, value) => {
                check(String(key));
                (p as unknown as Record<string, unknown>)[String(key)] = value;
                return true;
              },
            },
          ),
      );
    const range = {
      get text() {
        return state.text;
      },
      font: {
        get size() {
          return state.size;
        },
        set size(value: number) {
          check('size');
          state.size = value;
        },
        get bold() {
          return state.bold;
        },
        set bold(value: boolean | null) {
          check('bold');
          state.bold = value;
        },
        load: vi.fn(),
      },
      get paragraphs() {
        return { items: paragraphs(), load: vi.fn() };
      },
      load: vi.fn(),
      track: () => range,
      untrack: vi.fn(),
      compareLocationWith: () => ({ value: state.selected === id ? 'Equal' : 'Before' }),
      getRange: () => rangeFor(id),
      getOoxml: () => ({
        value: `<w:document><w:body>${state.paragraphs.map((p) => `<w:p><w:pPr><w:spacing w:line="${p.lineSpacing * 20}" w:lineRule="${p.rule}"/></w:pPr></w:p>`).join('')}</w:body></w:document>`,
      }),
    };
    return range;
  }
  const context = {
    document: { getSelection: () => rangeFor(state.selected) },
    sync: vi.fn(async () => undefined),
  };
  vi.stubGlobal('Office', { context: { document } });
  vi.stubGlobal('Word', {
    LocationRelation: { equal: 'Equal' },
    RangeLocation: { whole: 'Whole' },
    run: async (...args: unknown[]) => (args.at(-1) as (ctx: typeof context) => unknown)(context),
  });
  const host = new OfficeAdapter('text', true);
  return { state, document, host, runner: new TaskRunner(host, new FeatureRegistry(features)) };
}

afterEach(() => vi.unstubAllGlobals());

describe('Word 原生格式适配器（宿主替身）', () => {
  it.each([
    ['text.bold', {}, 'bold'],
    ['paragraph.center', {}, 'alignment'],
    ['paragraph.spacing', { lineSpacingMultiple: 2 }, 'lineSpacing'],
    ['paragraph.indent', { firstLineIndentPoints: 24 }, 'firstLineIndent'],
  ] as const)(
    '%s 仅写指定属性，按每段原值恢复，恢复不依赖当前选区',
    async (id, parameters, property) => {
      const { state, host, runner } = setup();
      const before = await host.readSelection();
      const preview = await runner.prepare(id, parameters);
      expect(state.writes).toEqual([]);
      await runner.execute(preview.id);
      expect(state.writes.every((p) => p === property)).toBe(true);
      const after = await host.readSelection();
      expect(after).not.toEqual(before);
      expect(after.text).toBe(before.text);
      expect(after.paragraphs!.map((p) => [p.text, p.leftIndent, p.rightIndent])).toEqual(
        before.paragraphs!.map((p) => [p.text, p.leftIndent, p.rightIndent]),
      );
      state.selected = 'elsewhere';
      await runner.undo();
      state.selected = 'original';
      const restored = await host.readSelection();
      // Selection IDs may change when the user returns, so compare content and formatting.
      expect(restored.text).toBe(before.text);
      expect(restored.bold).toBe(before.bold);
      expect(restored.paragraphs).toEqual(before.paragraphs);
    },
  );

  it.each(['text', 'alignment', 'lineSpacing', 'firstLineIndent'] as const)(
    '预览后完整段落的 %s 改变就拒绝写入',
    async (key) => {
      const { state, runner } = setup();
      const preview = await runner.prepare('paragraph.center');
      const p = state.paragraphs[0]!;
      if (key === 'text') p.text += '用户后续编辑';
      else if (key === 'alignment') p.alignment = 'Justified';
      else p[key] += 6;
      await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(state.writes).toEqual([]);
    },
  );

  it('后续编辑或换文档时拒绝恢复，且不写入', async () => {
    const { state, document, runner } = setup();
    await runner.execute((await runner.prepare('paragraph.center')).id);
    state.writes = [];
    state.paragraphs[1]!.text += '新文字';
    await expect(runner.undo()).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    state.paragraphs[1]!.text = '第二段的完整内容';
    document.url = 'file:///其他.docx';
    await expect(runner.undo()).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(state.writes).toEqual([]);
  });

  it('固定行距、编号缩进、混合粗细及超出20段均不写入', async () => {
    const { state, host, runner } = setup();
    state.paragraphs[0]!.rule = 'exact';
    await expect(
      runner.prepare('paragraph.spacing', { lineSpacingMultiple: 1.5 }),
    ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    state.paragraphs[0]!.isListItem = true;
    await expect(
      runner.prepare('paragraph.indent', { firstLineIndentPoints: 24 }),
    ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    state.bold = null;
    await expect(runner.prepare('text.bold')).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    state.paragraphs = Array.from({ length: 21 }, () => ({ ...state.paragraphs[0]! }));
    await expect(host.readSelection()).rejects.toMatchObject({ code: 'SELECTION_TOO_LARGE' });
    expect(state.writes).toEqual([]);
  });

  it('写入失败时没有成功结果或可恢复凭据', async () => {
    const { state, runner } = setup();
    const preview = await runner.prepare('text.bold');
    state.protected = true;
    await expect(runner.execute(preview.id)).rejects.toThrow('格式可能已修改');
    expect(runner.canUndo).toBe(false);
    expect(runner.history).toEqual([]);
  });
});
