import { describe, expect, it } from 'vitest';
import { WpsAdapter } from './adapter';
import { readWpsParagraphs } from './paragraphs';
import type { WpsApplication, WpsRange, WpsSelection } from './bridge';
import { createRuntime } from '../../app/runtime';

function setup(count = 3) {
  const document = { Name: '段落测试.docx', FullName: 'C:/test/段落测试.docx', ProtectionType: -1 };
  const parts = Array.from({ length: count }, (_, i) => ({
    text: i === 0 ? '前缀目标后缀\r' : `第${i + 1}段落\r`,
    alignment: i === 1 ? 31 : 17,
    listType: 55,
    spacingRule: 61,
    spacing: 12,
    indent: 0,
    characterIndent: 0 as number | undefined,
    formatWrites: [] as string[],
    writes: [] as number[],
    fail: false,
    ignore: false,
  }));
  const state = {
    start: 2,
    end: 4,
    size: 18,
    bold: 0 as number | boolean | undefined,
    fontWrites: [] as string[],
    story: 41,
    tables: 0,
  };
  const bounds = () => {
    let offset = 0;
    return parts.map((part) => {
      const start = offset;
      offset += part.text.length;
      return { start, end: offset };
    });
  };
  function rangeFor(start: number, end: number, partIndex?: number): WpsRange {
    const part = partIndex === undefined ? undefined : parts[partIndex]!;
    return {
      Start: start,
      End: end,
      get Text() {
        return parts
          .map((p) => p.text)
          .join('')
          .slice(start, end);
      },
      set Text(_value: string) {
        throw new Error('No text writes allowed');
      },
      Document: document,
      get Duplicate() {
        return rangeFor(start, end, partIndex);
      },
      get StoryType() {
        return state.story;
      },
      get Tables() {
        return { Count: state.tables };
      },
      get Paragraphs() {
        // Deliberately include the paragraph at exclusive End, as some native collections do.
        const selected = bounds()
          .map((b, i) => ({ ...b, i }))
          .filter((b) => b.start <= end && b.end > start);
        return {
          Count: selected.length,
          Item(index: number) {
            const b = selected[index - 1]!;
            return { Range: rangeFor(b.start, b.end, b.i) };
          },
        };
      },
      Font: {
        get Size() {
          return state.size;
        },
        set Size(value) {
          state.fontWrites.push('size');
          state.size = value;
          for (const p of parts) if (p.characterIndent) p.indent = p.characterIndent * value;
        },
        get Bold() {
          return state.bold;
        },
        set Bold(value) {
          state.fontWrites.push('bold');
          state.bold = value;
        },
      },
      ...(part
        ? {
            ListFormat: {
              get ListType() {
                return part.listType;
              },
            },
            ParagraphFormat: {
              get Alignment() {
                return part.alignment;
              },
              set Alignment(value: number | undefined) {
                part.writes.push(value!);
                if (!part.ignore) part.alignment = value!;
                if (part.fail) throw new Error('partial bridge write');
              },
              get LineSpacingRule() {
                return part.spacingRule;
              },
              set LineSpacingRule(value: number | undefined) {
                part.formatWrites.push('rule');
                if (!part.ignore) {
                  part.spacingRule = value!;
                  const index = [61, 67, 71].indexOf(value!);
                  if (index >= 0) part.spacing = [12, 18, 24][index]!;
                }
                if (part.fail) throw new Error('partial bridge write');
              },
              get LineSpacing() {
                return part.spacing;
              },
              set LineSpacing(value: number | undefined) {
                part.formatWrites.push('spacing');
                if (!part.ignore) part.spacing = value!;
                if (part.fail) throw new Error('partial bridge write');
              },
              get FirstLineIndent() {
                return part.indent;
              },
              set FirstLineIndent(value: number | undefined) {
                part.formatWrites.push('indent');
                if (!part.ignore && part.characterIndent === 0) part.indent = value!;
                if (part.fail) throw new Error('partial bridge write');
              },
              get CharacterUnitFirstLineIndent() {
                return part.characterIndent;
              },
              set CharacterUnitFirstLineIndent(value: number | undefined) {
                part.formatWrites.push('characters');
                if (!part.ignore) {
                  part.characterIndent = value!;
                  if (value) part.indent = value * 18;
                }
                if (part.fail) throw new Error('partial bridge write');
              },
              LeftIndent: 0,
              RightIndent: 0,
            },
          }
        : {}),
    };
  }
  const selection: WpsSelection = {
    get Start() {
      return state.start;
    },
    get End() {
      return state.end;
    },
    get Text() {
      return rangeFor(state.start, state.end).Text;
    },
    get Font() {
      return { Size: state.size, Bold: state.bold };
    },
    get Range() {
      return rangeFor(state.start, state.end);
    },
  };
  // Intentionally nonstandard enum numbers: production must use the host's runtime constants.
  const app: WpsApplication = {
    ActiveDocument: document,
    ActiveWindow: { Hwnd: 12 },
    Selection: selection,
    Enum: {
      WdParagraphAlignment: {
        wdAlignParagraphLeft: 17,
        wdAlignParagraphCenter: 23,
        wdAlignParagraphRight: 31,
        wdAlignParagraphJustify: 37,
      },
      WdStoryType: { wdMainTextStory: 41 },
      WdListType: { wdListNoNumbering: 55 },
      WdLineSpacing: {
        wdLineSpaceSingle: 61,
        wdLineSpace1pt5: 67,
        wdLineSpaceDouble: 71,
        wdLineSpaceMultiple: 79,
      },
    },
  };
  return { app, state, parts, bounds, rangeFor, host: new WpsAdapter('text', () => app) };
}

describe('WPS 完整段落居中与恢复', () => {
  it('兼容实机的扁平运行时枚举，不硬编码枚举数值', async () => {
    const { host, app, parts } = setup();
    app.Enum = {
      ...app.Enum!.WdParagraphAlignment,
      ...app.Enum!.WdStoryType,
      ...app.Enum!.WdListType,
      ...app.Enum!.WdLineSpacing,
    };
    const receipt = await host.apply({ type: 'paragraph.center' }, await host.readSelection());
    expect(parts[0]!.alignment).toBe(23);
    await host.undo(receipt);
    expect(parts[0]!.alignment).toBe(17);
  });
  it('部分选字预览完整段落，确认只改原生 Alignment，恢复原值', async () => {
    const { host, parts, state } = setup();
    const before = await host.readSelection();
    expect(before.text).toBe('目标');
    expect(before.paragraphs).toMatchObject([
      { text: '前缀目标后缀\r', alignment: 'Left', spacingMode: 'multiple' },
    ]);
    expect(parts.map((p) => p.writes)).toEqual([[], [], []]);
    const receipt = await host.apply({ type: 'paragraph.center' }, before);
    expect(parts.map((p) => p.alignment)).toEqual([23, 31, 17]);
    await host.undo(receipt);
    expect(parts.map((p) => p.alignment)).toEqual([17, 31, 17]);
    expect(parts.map((p) => p.writes)).toEqual([[23, 17], [], []]);
    expect(state.fontWrites).toEqual([]);
    expect(await host.readSelection()).toEqual(before);
  });

  it('多段不同原值逐段恢复，End 恰为下一段 Start 时不改下一段', async () => {
    const { host, state, bounds, parts } = setup();
    state.end = bounds()[1]!.end;
    const before = await host.readSelection();
    expect(before.paragraphs).toHaveLength(2);
    const receipt = await host.apply({ type: 'paragraph.center' }, before);
    expect(parts.map((p) => p.alignment)).toEqual([23, 23, 17]);
    state.start = bounds()[2]!.start;
    state.end = bounds()[2]!.end;
    await host.readSelection();
    await host.undo(receipt);
    expect(parts.map((p) => p.alignment)).toEqual([17, 31, 17]);
    expect(parts[2]!.writes).toEqual([]);
  });

  it.each(['text', 'alignment', 'spacingRule', 'indent'] as const)(
    '同段未选中部分或格式 %s 变化使旧预览失效',
    async (field) => {
      const { host, parts } = setup();
      const before = await host.readSelection();
      if (field === 'text') parts[0]!.text = '新的目标内容\r';
      else if (field === 'alignment') parts[0]!.alignment = 31;
      else if (field === 'spacingRule') parts[0]!.spacingRule = 2;
      else parts[0]!.indent = 24;
      await expect(host.apply({ type: 'paragraph.center' }, before)).rejects.toMatchObject({
        code: 'STALE_SELECTION',
      });
      expect(parts[0]!.writes).toEqual([]);
    },
  );

  it.each(['text', 'alignment', 'spacingRule', 'indent'] as const)(
    '居中后 %s 变化时恢复不覆盖用户编辑',
    async (field) => {
      const { host, parts } = setup();
      const receipt = await host.apply({ type: 'paragraph.center' }, await host.readSelection());
      if (field === 'text') parts[0]!.text = '新的目标内容\r';
      else if (field === 'alignment') parts[0]!.alignment = 31;
      else if (field === 'spacingRule') parts[0]!.spacingRule = 2;
      else parts[0]!.indent = 24;
      await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(parts[0]!.writes).toEqual([23]);
    },
  );

  it.each(['enum', 'unknown', 'list', 'table', 'story', 'spacing'] as const)(
    '无法可靠恢复 %s 时不提供段落状态，既有字号仍能用',
    async (unsupported) => {
      const { host, app, parts, state } = setup();
      if (unsupported === 'enum') delete app.Enum;
      else if (unsupported === 'unknown') parts[0]!.alignment = 999;
      else if (unsupported === 'list') parts[0]!.listType = 56;
      else if (unsupported === 'table') state.tables = 1;
      else if (unsupported === 'story') state.story = 99;
      else parts[0]!.spacing = NaN;
      const before = await host.readSelection();
      expect(before.paragraphs).toBeUndefined();
      await expect(host.apply({ type: 'paragraph.center' }, before)).rejects.toMatchObject({
        code: 'UNAVAILABLE',
      });
      const receipt = await host.apply({ type: 'text.resize', size: 22 }, before);
      await host.undo(receipt);
      expect(state.size).toBe(18);
      expect(parts[0]!.writes).toEqual([]);
    },
  );

  it('完整段落超过20,000字符时拒绝居中，部分选字的字号仍可用', async () => {
    const { host, parts } = setup();
    parts[0]!.text = '前缀目标' + '长'.repeat(20000) + '\r';
    const before = await host.readSelection();
    expect(before.paragraphs).toBeUndefined();
    await expect(host.apply({ type: 'paragraph.center' }, before)).rejects.toMatchObject({
      code: 'UNAVAILABLE',
    });
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, before);
    await host.undo(receipt);
  });

  it('允许20段和尾部独占边界，拒绝21个实际相交段落', async () => {
    const { host, state, bounds } = setup(21);
    state.end = bounds()[19]!.end;
    expect((await host.readSelection()).paragraphs).toHaveLength(20);
    state.end += 1;
    expect((await host.readSelection()).paragraphs).toBeUndefined();
  });

  it('缺失桌面接口时只禁用居中', () => {
    const { app, rangeFor } = setup();
    const range = rangeFor(2, 4);
    Object.defineProperty(range, 'Paragraphs', {
      get() {
        throw new Error('unsupported');
      },
    });
    expect(readWpsParagraphs(range, app)).toBeUndefined();
  });

  it.each(['fail', 'ignore'] as const)('多段部分写入 %s 后清除旧恢复凭证', async (failure) => {
    const { host, state, bounds, parts } = setup();
    const previous = await host.apply(
      { type: 'text.resize', size: 22 },
      await host.readSelection(),
    );
    state.end = bounds()[1]!.end;
    parts[1]![failure] = true;
    await expect(
      host.apply({ type: 'paragraph.center' }, await host.readSelection()),
    ).rejects.toMatchObject({ code: 'INVALID_PLAN' });
    await expect(host.undo(previous)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(parts[0]!.alignment).toBe(23);
  });

  it('居中后仍可修改字号并恢复，段落原值不受影响', async () => {
    const { host, parts, state } = setup();
    await host.apply({ type: 'paragraph.center' }, await host.readSelection());
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, await host.readSelection());
    await host.undo(receipt);
    expect(state.size).toBe(18);
    expect(parts[0]!.alignment).toBe(23);
  });

  it('段落回读不可用时不发恢复凭证，也不能覆盖之后的对齐', async () => {
    const { host, app, parts } = setup();
    const receipt = await host.apply({ type: 'paragraph.center' }, await host.readSelection());
    delete app.Enum;
    parts[0]!.alignment = 31;
    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(parts[0]!.writes).toEqual([23]);
  });
});

describe('WPS 行距与首行缩进', () => {
  it('原有字符缩进随字号变化，字号恢复仍保留原始单位和磅值', async () => {
    const { host, parts } = setup();
    parts[0]!.characterIndent = 2;
    parts[0]!.indent = 36;
    const before = await host.readSelection();
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, before);
    expect(parts[0]!.indent).toBe(44);
    await host.undo(receipt);
    expect(await host.readSelection()).toEqual(before);
    expect(parts[0]!.formatWrites).toEqual([]);
  });
  it.each([1, 1.5, 2, 1.2, 3])(
    '支持 %s 倍行距，预览取消不写入、恢复保留逐段规则',
    async (multiple) => {
      const { host, state, bounds, parts } = setup();
      parts[1]!.spacingRule = 79;
      parts[1]!.spacing = 30;
      state.end = bounds()[1]!.end;
      const original = await host.readSelection();
      const { runner } = createRuntime(host);
      const canceled = await runner.prepare('paragraph.spacing', { lineSpacingMultiple: multiple });
      expect(canceled.summary).toContain('2 个完整段落');
      runner.cancel(canceled.id);
      expect(parts.flatMap((p) => p.formatWrites)).toEqual([]);
      const task = await runner.prepare('paragraph.spacing', { lineSpacingMultiple: multiple });
      await runner.execute(task.id);
      expect((await host.readSelection()).paragraphs!.map((p) => p.lineSpacing)).toEqual([
        Math.round(multiple * 240) / 20,
        Math.round(multiple * 240) / 20,
      ]);
      expect(parts[2]!.formatWrites).toEqual([]);
      state.start = bounds()[2]!.start;
      state.end = bounds()[2]!.end;
      await runner.undo();
      state.start = 2;
      state.end = bounds()[1]!.end;
      expect(await host.readSelection()).toEqual(original);
      expect(parts.map((p) => p.spacingRule)).toEqual([61, 79, 61]);
      expect(parts.map((p) => p.spacing)).toEqual([12, 30, 12]);
      expect(state.fontWrites).toEqual([]);
      expect(parts.map((p) => p.alignment)).toEqual([17, 31, 17]);
    },
  );

  it.each([0, 24, 28.35, 144])('首行缩进 %s 磅兼容原始字符/磅缩进并逐段恢复', async (points) => {
    const { host, state, bounds, parts } = setup();
    parts[0]!.characterIndent = 2;
    parts[0]!.indent = 36;
    parts[1]!.indent = -12;
    state.end = bounds()[1]!.end;
    const original = await host.readSelection();
    const receipt = await host.apply({ type: 'paragraph.indent', points }, original);
    expect(parts.map((p) => p.indent)).toEqual([points, points, 0]);
    expect(parts.map((p) => p.characterIndent)).toEqual([0, 0, 0]);
    expect(parts[0]!.formatWrites).toEqual(['characters', 'indent']);
    expect(parts[2]!.formatWrites).toEqual([]);
    await host.undo(receipt);
    expect(await host.readSelection()).toEqual(original);
    expect(parts[0]!.formatWrites).toEqual([
      'characters',
      'indent',
      'characters',
      'indent',
      'characters',
    ]);
    expect(parts.map((p) => p.spacing)).toEqual([12, 12, 12]);
    expect(state.fontWrites).toEqual([]);
  });

  it.each(['paragraph.spacing', 'paragraph.indent'] as const)(
    '%s 拒绝过期预览和覆盖后续字符缩进',
    async (type) => {
      const { host, parts } = setup();
      const plan = type === 'paragraph.spacing' ? { type, multiple: 1.2 } : { type, points: 24 };
      const before = await host.readSelection();
      parts[0]!.characterIndent = 1;
      await expect(host.apply(plan, before)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(parts[0]!.formatWrites).toEqual([]);
      parts[0]!.characterIndent = 0;
      const receipt = await host.apply(plan, await host.readSelection());
      parts[0]!.characterIndent = 1;
      const writes = [...parts[0]!.formatWrites];
      await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
      expect(parts[0]!.formatWrites).toEqual(writes);
    },
  );

  it.each(['enum', 'fixed', 'canonicalMultiple', 'inconsistentPreset'] as const)(
    '不安全行距 %s 在预览前拒绝，居中和缩进可继续',
    async (kind) => {
      const { host, parts, app } = setup();
      if (kind === 'enum') delete app.Enum!.WdLineSpacing;
      else if (kind === 'fixed') parts[0]!.spacingRule = 88;
      else if (kind === 'canonicalMultiple') {
        parts[0]!.spacingRule = 79;
        parts[0]!.spacing = 18;
      } else parts[0]!.spacing = 30;
      const { runner } = createRuntime(host);
      await expect(
        runner.prepare('paragraph.spacing', { lineSpacingMultiple: 1.5 }),
      ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
      expect(parts[0]!.formatWrites).toEqual([]);
      for (const plan of [
        { type: 'paragraph.center' },
        { type: 'paragraph.indent', points: 24 },
      ] as const) {
        const receipt = await host.apply(plan, await host.readSelection());
        await host.undo(receipt);
      }
    },
  );

  it('缺少字符缩进接口仅阻止缩进预览', async () => {
    const { host, parts } = setup();
    parts[0]!.characterIndent = undefined;
    const { runner } = createRuntime(host);
    await expect(
      runner.prepare('paragraph.indent', { firstLineIndentPoints: 24 }),
    ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    expect(parts[0]!.formatWrites).toEqual([]);
    for (const plan of [
      { type: 'paragraph.center' },
      { type: 'paragraph.spacing', multiple: 1.2 },
    ] as const) {
      const receipt = await host.apply(plan, await host.readSelection());
      await host.undo(receipt);
    }
  });

  it.each(['paragraph.spacing', 'paragraph.indent'] as const)(
    '%s 多段部分写入失败清除旧凭证，回读不符不给凭证',
    async (type) => {
      for (const failure of ['fail', 'ignore'] as const) {
        const { host, state, bounds, parts } = setup();
        const old = await host.apply({ type: 'text.bold', bold: true }, await host.readSelection());
        state.end = bounds()[1]!.end;
        parts[1]![failure] = true;
        const plan = type === 'paragraph.spacing' ? { type, multiple: 1.2 } : { type, points: 24 };
        await expect(host.apply(plan, await host.readSelection())).rejects.toMatchObject({
          code: 'INVALID_PLAN',
        });
        await expect(host.undo(old)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
      }
    },
  );
});
