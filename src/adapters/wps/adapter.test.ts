import { describe, expect, it } from 'vitest';
import { WpsAdapter } from './adapter';
import type { WpsApplication, WpsDocument, WpsRange, WpsSelection } from './bridge';

interface TextState {
  text: string;
  start: number;
  end: number;
  size: number;
  writes: number[];
}

function setup() {
  const document: WpsDocument = {
    Name: '通知.docx',
    FullName: 'C:/test/通知.docx',
    ReadOnly: false,
    ProtectionType: -1,
  };
  const original: TextState = { text: '活动通知', start: 0, end: 4, size: 18, writes: [] };
  const other: TextState = { text: '其他内容', start: 10, end: 14, size: 12, writes: [] };

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
        return { Size: state.size };
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
