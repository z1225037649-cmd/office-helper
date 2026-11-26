import { describe, expect, it } from 'vitest';
import { WpsAdapter } from './adapter';
import type { WpsApplication, WpsDocument, WpsRange, WpsSelection } from './bridge';

interface TextState {
  text: string;
  start: number;
  end: number;
  size: number;
  writes: number[];
  bold?: number | boolean;
  boldWrites: (number | boolean | undefined)[];
}

function setup() {
  const document: WpsDocument = {
    Name: '通知.docx',
    FullName: 'C:/test/通知.docx',
    ReadOnly: false,
    ProtectionType: -1,
  };
  const original: TextState = {
    text: '活动通知',
    start: 0,
    end: 4,
    size: 18,
    writes: [],
    bold: 0,
    boldWrites: [],
  };
  const other: TextState = {
    text: '其他内容',
    start: 10,
    end: 14,
    size: 12,
    writes: [],
    bold: 0,
    boldWrites: [],
  };

  // 每个句柄绑定创建时的文字区域，移动 Selection 不会重定向已经取得的 Duplicate。
  function rangeFor(state: TextState): WpsRange {
    return {
      get Text() {
        return state.text;
      },
      set Text(text) {
        state.text = text;
      },
      get Start() {
        return state.start;
      },
      set Start(start) {
        state.start = start;
      },
      get End() {
        return state.end;
      },
      set End(end) {
        state.end = end;
      },
      Font: {
        get Bold() {
          return state.bold;
        },
        set Bold(bold) {
          state.boldWrites.push(bold);
          state.bold = bold;
        },
        get Size() {
          return state.size;
        },
        set Size(size) {
          state.writes.push(size);
          state.size = size;
        },
      },
      Document: document,
      get Duplicate() {
        return rangeFor(state);
      },
    };
  }

  function selectionFor(state: TextState): WpsSelection {
    return {
      get Text() {
        return state.text;
      },
      get Start() {
        return state.start;
      },
      get End() {
        return state.end;
      },
      get Font() {
        return { Size: state.size, Bold: state.bold };
      },
      get Range() {
        return rangeFor(state);
      },
    };
  }

  const app: WpsApplication = {
    ActiveDocument: document,
    ActiveWindow: { Hwnd: 101 },
    Selection: selectionFor(original),
  };
  const host = new WpsAdapter('text', () => app);
  return { host, app, document, original, other, rangeFor, selectionFor };
}

describe('WPS 字号修改与恢复的安全边界', () => {
  it('移动选区并重新读取后，仍只恢复原始区域', async () => {
    const { host, app, original, other, selectionFor } = setup();
    const before = await host.readSelection();
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, before);
    expect(receipt.canUndo).toBe(true);
    expect(original.size).toBe(22);

    app.Selection = selectionFor(other);
    await host.readSelection();
    await host.undo(receipt);

    expect(original).toMatchObject({ text: before.text, size: 18, writes: [22, 18] });
    expect(other).toMatchObject({ size: 12, writes: [] });
    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
  });

  it('位置改变时即使文字与字号相同，也拒绝旧预览', async () => {
    const { host, app, original, other, selectionFor } = setup();
    const before = await host.readSelection();
    other.text = original.text;
    other.size = original.size;
    app.Selection = selectionFor(other);

    await expect(host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(original.writes).toEqual([]);
    expect(other.writes).toEqual([]);
  });

  it.each(['text', 'size'] as const)('预览后 %s 改变时拒绝写入', async (field) => {
    const { host, original } = setup();
    const before = await host.readSelection();
    if (field === 'text') original.text = '最新通知';
    else original.size = 26;

    await expect(host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(original.writes).toEqual([]);
    expect(original[field]).toBe(field === 'text' ? '最新通知' : 26);
  });

  it('换文件后即使选区内容相同，也拒绝旧预览', async () => {
    const { host, app, original } = setup();
    const before = await host.readSelection();
    app.ActiveDocument = { Name: '另一份.docx', FullName: 'C:/test/另一份.docx' };

    await expect(host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(original.writes).toEqual([]);
  });

  it('核对快照后、获取写入句柄前切换文件时，也不能写入另一份同内容文档', async () => {
    const first = setup();
    const second = setup();
    second.document.Name = '另一份.docx';
    second.document.FullName = 'C:/test/另一份.docx';
    let applicationReads = 0;
    const host = new WpsAdapter('text', () => {
      applicationReads += 1;
      return applicationReads <= 2 ? first.app : second.app;
    });
    const before = await host.readSelection();

    await expect(host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(first.original.writes).toEqual([]);
    expect(second.original.writes).toEqual([]);
  });

  it('独立 Range 与刚读到的选区不一致时不写入', async () => {
    const { host, app, original, other, selectionFor, rangeFor } = setup();
    const before = await host.readSelection();
    app.Selection = { ...selectionFor(original), Range: rangeFor(other) };

    await expect(host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(original.writes).toEqual([]);
    expect(other.writes).toEqual([]);
  });

  it('写入句柄必须属于已经核对的文档', async () => {
    const first = setup();
    const second = setup();
    second.document.Name = '另一份.docx';
    second.document.FullName = 'C:/test/另一份.docx';
    const before = await first.host.readSelection();
    first.app.Selection = {
      ...first.selectionFor(first.original),
      Range: second.rangeFor(second.original),
    };

    await expect(first.host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject(
      { code: 'STALE_SELECTION' },
    );
    expect(first.original.writes).toEqual([]);
    expect(second.original.writes).toEqual([]);
  });

  it.each([
    { ReadOnly: true, ProtectionType: -1 },
    { ReadOnly: false, ProtectionType: 0 },
    { ReadOnly: false, ProtectionType: 3 },
  ])('拒绝修改只读或受保护文档：%j', async (protection) => {
    const { host, document, original } = setup();
    const before = await host.readSelection();
    Object.assign(document, protection);

    await expect(host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject({
      code: 'UNAVAILABLE',
    });
    expect(original).toMatchObject({ size: 18, writes: [] });
  });

  it.each([
    { ReadOnly: true, ProtectionType: -1 },
    { ReadOnly: false, ProtectionType: 0 },
  ])('文档修改后被保护时拒绝恢复：%j', async (protection) => {
    const { host, document, original } = setup();
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, await host.readSelection());
    Object.assign(document, protection);

    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    expect(original).toMatchObject({ size: 22, writes: [22] });
  });

  it.each(['text', 'size'] as const)('恢复不会覆盖目标区域后续的 %s 修改', async (field) => {
    const { host, original } = setup();
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, await host.readSelection());
    if (field === 'text') original.text = '用户新改的通知';
    else original.size = 30;

    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(original.writes).toEqual([22]);
    expect(original[field]).toBe(field === 'text' ? '用户新改的通知' : 30);
  });

  it('在其他文件拒绝恢复，回到原文件后允许恢复', async () => {
    const { host, app, document, original } = setup();
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, await host.readSelection());
    app.ActiveDocument = { Name: '另一份.docx', FullName: 'C:/test/另一份.docx' };
    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(original.writes).toEqual([22]);

    app.ActiveDocument = document;
    await host.undo(receipt);
    expect(original).toMatchObject({ size: 18, writes: [22, 18] });
  });

  it('新的修改使旧恢复凭证失效，最新凭证只恢复最近一步', async () => {
    const { host, original } = setup();
    const first = await host.apply({ type: 'text.resize', size: 22 }, await host.readSelection());
    const second = await host.apply({ type: 'text.resize', size: 26 }, await host.readSelection());

    await expect(host.undo(first)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(original.size).toBe(26);
    await host.undo(second);
    expect(original).toMatchObject({ size: 22, writes: [22, 26, 22] });
  });

  it('关闭适配器后，旧预览与恢复凭证均不能继续操作', async () => {
    const { host, original } = setup();
    const before = await host.readSelection();
    const receipt = await host.apply({ type: 'text.resize', size: 22 }, before);
    host.dispose();

    await expect(host.readSelection()).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    await expect(host.apply({ type: 'text.resize', size: 26 }, before)).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(original).toMatchObject({ size: 22, writes: [22] });
  });
});

describe('WPS 加粗与取消加粗', () => {
  it.each([
    [0, false],
    [-1, true],
    [false, false],
    [true, true],
    [9999999, undefined],
    [1, undefined],
    [undefined, undefined],
  ])('仅将明确的原生粗细 %s 转为 %s', async (raw, expected) => {
    const { host, original } = setup();
    original.bold = raw as number | boolean | undefined;
    expect((await host.readSelection()).bold).toBe(expected);
  });

  it.each([false, true])('支持改为 %s 并恢复，只写 Bold', async (bold) => {
    const { host, original } = setup();
    original.bold = bold ? 0 : -1;
    const before = await host.readSelection();
    const receipt = await host.apply({ type: 'text.bold', bold }, before);
    expect((await host.readSelection()).bold).toBe(bold);
    await host.undo(receipt);
    expect((await host.readSelection()).bold).toBe(!bold);
    expect(original.boldWrites).toEqual(bold ? [-1, 0] : [0, -1]);
    expect(original.writes).toEqual([]);
    expect(original.text).toBe(before.text);
  });

  it('加粗不需要统一字号，也不写字号', async () => {
    const { host, original } = setup();
    original.size = 9999999;
    const receipt = await host.apply({ type: 'text.bold', bold: true }, await host.readSelection());
    await host.undo(receipt);
    expect(original.size).toBe(9999999);
    expect(original.writes).toEqual([]);
  });

  it.each([9999999, undefined])('粗细未知或混合 %s 时拒绝加粗', async (bold) => {
    const { host, original } = setup();
    original.bold = bold;
    await expect(
      host.apply({ type: 'text.bold', bold: true }, await host.readSelection()),
    ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
    expect(original.boldWrites).toEqual([]);
  });

  it('预览后的粗细变化使旧字号和加粗计划都失效', async () => {
    const { host, original } = setup();
    const before = await host.readSelection();
    original.bold = -1;
    for (const plan of [
      { type: 'text.bold', bold: true },
      { type: 'text.resize', size: 22 },
    ] as const)
      await expect(host.apply(plan, before)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(original.boldWrites).toEqual([]);
    expect(original.writes).toEqual([]);
  });

  it('获取的独立 Range 粗细不同于快照时不写入', async () => {
    const { host, app, original, rangeFor, selectionFor } = setup();
    const before = await host.readSelection();
    const range = rangeFor(original);
    range.Font = { Size: 18, Bold: -1 };
    app.Selection = { ...selectionFor(original), Range: range };
    Object.defineProperty(range, 'Duplicate', { value: range });
    await expect(host.apply({ type: 'text.bold', bold: true }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    expect(original.boldWrites).toEqual([]);
  });

  it('移动选区不影响加粗恢复位置', async () => {
    const { host, app, original, other, selectionFor } = setup();
    const receipt = await host.apply({ type: 'text.bold', bold: true }, await host.readSelection());
    app.Selection = selectionFor(other);
    await host.readSelection();
    await host.undo(receipt);
    expect(original.boldWrites).toEqual([-1, 0]);
    expect(other.boldWrites).toEqual([]);
  });

  it.each(['text', 'size', 'bold'] as const)('恢复加粗不会覆盖后来更改的 %s', async (field) => {
    const { host, original } = setup();
    const receipt = await host.apply({ type: 'text.bold', bold: true }, await host.readSelection());
    if (field === 'text') original.text = '新的正文';
    else if (field === 'size') original.size = 30;
    else original.bold = 0;
    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(original.boldWrites).toEqual([-1]);
  });

  it('换文件拒绝执行与恢复加粗', async () => {
    const { host, app, document, original } = setup();
    const before = await host.readSelection();
    const receipt = await host.apply({ type: 'text.bold', bold: true }, before);
    app.ActiveDocument = { Name: '另一份.docx', FullName: 'C:/test/另一份.docx' };
    await expect(host.undo(receipt)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    await expect(host.apply({ type: 'text.bold', bold: false }, before)).rejects.toMatchObject({
      code: 'STALE_SELECTION',
    });
    app.ActiveDocument = document;
    await host.undo(receipt);
    expect(original.boldWrites).toEqual([-1, 0]);
  });

  it.each([{ ReadOnly: true }, { ProtectionType: 0 }])(
    '加粗尊重文档保护 %j',
    async (protection) => {
      const { host, document, original } = setup();
      Object.assign(document, protection);
      await expect(
        host.apply({ type: 'text.bold', bold: true }, await host.readSelection()),
      ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
      expect(original.boldWrites).toEqual([]);
    },
  );

  it('无法读取完整段落时拒绝段落格式操作', async () => {
    const { host, original } = setup();
    expect(host.info.capabilities).toContain('text.bold');
    expect(host.info.capabilities).toContain('paragraph.center');
    for (const plan of [
      { type: 'paragraph.center' },
      { type: 'paragraph.spacing', multiple: 1.5 },
      { type: 'paragraph.indent', points: 24 },
    ] as const)
      await expect(host.apply(plan, await host.readSelection())).rejects.toMatchObject({
        code: 'UNAVAILABLE',
      });
    expect(original.boldWrites).toEqual([]);
  });

  it.each(['text.resize', 'text.bold'] as const)(
    '插入光标即使返回右侧文字也拒绝 %s',
    async (type) => {
      const { host, original } = setup();
      original.end = original.start;
      const snapshot = await host.readSelection();
      expect(snapshot.text).toBe('');
      await expect(
        host.apply(type === 'text.resize' ? { type, size: 22 } : { type, bold: true }, snapshot),
      ).rejects.toMatchObject({ code: 'EMPTY_SELECTION' });
      expect(original.writes).toEqual([]);
      expect(original.boldWrites).toEqual([]);
    },
  );

  it('部分写入失败不能留下旧恢复凭证', async () => {
    const { host, app, original, rangeFor, selectionFor } = setup();
    const previous = await host.apply(
      { type: 'text.resize', size: 22 },
      await host.readSelection(),
    );
    const range = rangeFor(original);
    Object.defineProperty(range.Font, 'Bold', {
      get: () => original.bold,
      set: (value) => {
        original.bold = value;
        throw new Error('bridge failure');
      },
    });
    Object.defineProperty(range, 'Duplicate', { value: range });
    app.Selection = { ...selectionFor(original), Range: range };
    await expect(
      host.apply({ type: 'text.bold', bold: true }, await host.readSelection()),
    ).rejects.toMatchObject({ code: 'INVALID_PLAN' });
    await expect(host.undo(previous)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(original.bold).toBe(-1);
  });
});
