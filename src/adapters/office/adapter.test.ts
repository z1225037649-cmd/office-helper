import { afterEach, describe, expect, it, vi } from 'vitest';
import { OfficeAdapter } from './adapter';
import { FeatureRegistry } from '../../core/registry';
import { TaskRunner } from '../../core/runner';
import { features } from '../../features';

function setupReadOnlyWord() {
  const state = { text: '活动通知', size: 18 };
  const document = { url: 'https://example.test/通知.docx' };
  const compareLocationWith = vi.fn(() => {
    throw new Error('WordApi 1.3 unavailable');
  });
  const track = vi.fn();
  const untrack = vi.fn();
  const range = {
    get text() {
      return state.text;
    },
    font: {
      get size() {
        return state.size;
      },
      load: vi.fn(),
    },
    load: vi.fn(),
    compareLocationWith,
    track,
    untrack,
  };
  track.mockReturnValue(range);
  const context = { document: { getSelection: () => range }, sync: vi.fn(async () => undefined) };
  const run = vi.fn(async (...args: unknown[]) => {
    const callback = args.at(-1) as (value: typeof context) => Promise<unknown>;
    return callback(context);
  });
  vi.stubGlobal('Office', { context: { document } });
  vi.stubGlobal('Word', { run });

  const host = new OfficeAdapter('text', false);
  const runner = new TaskRunner(host, new FeatureRegistry(features));
  return { host, runner, state, document, compareLocationWith, track, untrack };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('旧版 Word 只读选区快照', () => {
  it('重复读取同一内容时身份稳定，且不依赖比较位置或跟踪句柄', async () => {
    const { host, compareLocationWith, track, untrack } = setupReadOnlyWord();
    const first = await host.readSelection();
    const second = await host.readSelection();

    expect(second).toEqual(first);
    expect(first).toMatchObject({ text: '活动通知', fontSize: 18 });
    host.dispose();
    expect(compareLocationWith).not.toHaveBeenCalled();
    expect(track).not.toHaveBeenCalled();
    expect(untrack).not.toHaveBeenCalled();
  });

  it('查看选中内容能够从预览执行到结果，不被自身第二次读取判为过期', async () => {
    const { runner } = setupReadOnlyWord();
    const preview = await runner.prepare('selection.inspect');

    await expect(runner.execute(preview.id)).resolves.toMatchObject({
      message: '当前选中 4 个字符，字号 18 磅。',
      canUndo: false,
    });
  });

  it.each(['text', 'size'] as const)('预览后 %s 改变时仍拒绝过期报告', async (field) => {
    const { runner, state } = setupReadOnlyWord();
    const preview = await runner.prepare('selection.inspect');
    if (field === 'text') state.text = '不同的文字';
    else state.size = 30;

    await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(runner.history).toEqual([]);
  });

  it('预览后文件 URL 改变时拒绝过期报告', async () => {
    const { runner, document } = setupReadOnlyWord();
    const preview = await runner.prepare('selection.inspect');
    document.url = 'https://example.test/其他文件.docx';

    await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
  });

  it('稳定的只读快照不会开放字号写入能力', async () => {
    const { host, state } = setupReadOnlyWord();
    const before = await host.readSelection();
    expect(host.info.capabilities).not.toContain('text.resize');
    expect(host.info.capabilities).not.toContain('history.undo');

    await expect(host.apply({ type: 'text.resize', size: 22 }, before)).rejects.toMatchObject({
      code: 'UNAVAILABLE',
    });
    expect(state.size).toBe(18);
  });
});
