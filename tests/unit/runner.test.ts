import { describe, expect, it, vi } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import { WpsAdapter } from '../../src/adapters/wps/adapter';
import type { ActionPlan, FeatureDefinition } from '../../src/core/contracts';
import { FeatureRegistry } from '../../src/core/registry';
import { TaskRunner } from '../../src/core/runner';
import { features } from '../../src/features';

const setup = () => {
  const host = new MockAdapter();
  return { host, runner: new TaskRunner(host, new FeatureRegistry(features)) };
};

describe('任务执行边界', () => {
  it('预览不会修改文件，确认执行后可恢复原始字号', async () => {
    const { host, runner } = setup();
    const before = await host.readSelection();
    const preview = await runner.prepare('text.enlarge');
    expect((await host.readSelection()).fontSize).toBe(before.fontSize);
    expect(preview.changesDocument).toBe(true);
    await runner.execute(preview.id);
    expect((await host.readSelection()).fontSize).toBe(22);
    expect(runner.canUndo).toBe(true);
    await runner.undo();
    expect((await host.readSelection()).fontSize).toBe(18);
    expect(runner.canUndo).toBe(false);
  });
  it('明确字号会进入预览并按指定值执行', async () => {
    const { host, runner } = setup();
    const preview = await runner.prepare('text.enlarge', { fontSize: 30 });
    expect(preview.summary).toContain('30 磅');
    expect((await host.readSelection()).fontSize).toBe(18);
    await runner.execute(preview.id);
    expect((await host.readSelection()).fontSize).toBe(30);
  });

  it('拒绝已取消或重复执行的预览', async () => {
    const { runner } = setup();
    const first = await runner.prepare('text.enlarge');
    runner.cancel(first.id);
    await expect(runner.execute(first.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    const second = await runner.prepare('text.enlarge');
    await runner.execute(second.id);
    await expect(runner.execute(second.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
  });

  it('生成新预览会使旧预览失效', async () => {
    const { runner } = setup();
    const first = await runner.prepare('text.enlarge');
    await runner.prepare('selection.inspect');
    await expect(runner.execute(first.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
  });

  it('预览后文档改变时拒绝写入', async () => {
    const { host, runner } = setup();
    const preview = await runner.prepare('text.enlarge');
    await host.apply({ type: 'text.resize', size: 30 }, await host.readSelection());
    await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect((await host.readSelection()).fontSize).toBe(30);
  });

  it('修改后发生其他编辑时拒绝恢复，避免覆盖新编辑', async () => {
    const { host, runner } = setup();
    await runner.execute((await runner.prepare('text.enlarge')).id);
    await host.apply({ type: 'text.resize', size: 36 }, await host.readSelection());
    await expect(runner.undo()).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect((await host.readSelection()).fontSize).toBe(36);
  });

  it('读取信息不会清掉上次可恢复修改', async () => {
    const { runner } = setup();
    await runner.execute((await runner.prepare('text.enlarge')).id);
    await runner.execute((await runner.prepare('selection.inspect')).id);
    expect(runner.canUndo).toBe(true);
  });

  it('执行互斥，连续点击不会并发读取和写入', async () => {
    const { host, runner } = setup();
    const original = host.readSelection.bind(host);
    let release!: () => void;
    vi.spyOn(host, 'readSelection').mockImplementationOnce(async () => {
      await new Promise<void>((resolve) => {
        release = resolve;
      });
      return original();
    });
    const first = runner.prepare('text.enlarge');
    await expect(runner.prepare('text.enlarge')).rejects.toMatchObject({ code: 'BUSY' });
    release();
    await first;
  });

  it('功能不能通过漏报能力需求绕过宿主的写入限制', async () => {
    const hostile: FeatureDefinition = {
      ...features[0]!,
      requires: ['selection.read'],
      prepare: () => ({ summary: '修改', plan: { type: 'text.resize', size: 22 } }),
    };
    const host = new WpsAdapter('sheet', () => ({
      ActiveWorkbook: { Name: 'example.xlsx' },
      Selection: { Rows: { Count: 1 }, Columns: { Count: 1 }, Address: 'A1', Value2: 1 },
    }));
    hostile.kinds = ['sheet'];
    const runner = new TaskRunner(host, new FeatureRegistry([hostile]));
    await expect(runner.prepare(hostile.id)).rejects.toMatchObject({ code: 'UNAVAILABLE' });
  });

  it.each([NaN, Infinity, 0, 200])('拒绝无效字号 %s', async (size) => {
    const invalid = {
      ...features[0]!,
      prepare: () => ({ summary: '修改', plan: { type: 'text.resize', size } as ActionPlan }),
    };
    const runner = new TaskRunner(new MockAdapter(), new FeatureRegistry([invalid]));
    await expect(runner.prepare(invalid.id)).rejects.toMatchObject({ code: 'INVALID_PLAN' });
  });

  it('快照返回值与宿主内部数据隔离', async () => {
    const { host } = setup();
    const snapshot = await host.readSelection();
    snapshot.text = '外部改动';
    expect((await host.readSelection()).text).toBe('社区活动通知');
  });
});
