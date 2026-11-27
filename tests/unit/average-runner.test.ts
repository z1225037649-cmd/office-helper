import { describe, expect, it, vi } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import type { SelectionSnapshot } from '../../src/core/contracts';
import { FeatureRegistry } from '../../src/core/registry';
import { TaskRunner } from '../../src/core/runner';
import { features } from '../../src/features';

describe('平均值查询的执行与选区保护', () => {
  it('按需读取类型，预览及执行均不写文件，也不提供恢复', async () => {
    const host = new MockAdapter('sheet');
    const before = await host.readSelection();
    const read = vi.spyOn(host, 'readSelection');
    const apply = vi.spyOn(host, 'apply');
    const undo = vi.spyOn(host, 'undo');
    const runner = new TaskRunner(host, new FeatureRegistry(features));

    const preview = await runner.prepare('sheet.average');
    expect(preview.changesDocument).toBe(false);
    expect(preview.target).toContain('B2:B5');
    expect(read).toHaveBeenLastCalledWith({ includeCellTypes: true });
    const outcome = await runner.execute(preview.id);
    expect(outcome.message).toContain('112.625');
    expect(outcome.canUndo).toBe(false);
    expect(read).toHaveBeenCalledTimes(2);
    expect(read).toHaveBeenLastCalledWith({ includeCellTypes: true });
    expect(await host.readSelection()).toEqual(before);
    expect(apply).not.toHaveBeenCalled();
    expect(undo).not.toHaveBeenCalled();
    expect(runner.history).toHaveLength(1);
    expect(runner.history[0]?.featureId).toBe('sheet.average');
  });

  it('取消、重复确认和新预览使旧平均值任务失效', async () => {
    const host = new MockAdapter('sheet');
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    const cancelled = await runner.prepare('sheet.average');
    runner.cancel(cancelled.id);
    await expect(runner.execute(cancelled.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });

    const replaced = await runner.prepare('sheet.average');
    const current = await runner.prepare('sheet.average');
    await expect(runner.execute(replaced.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    await runner.execute(current.id);
    await expect(runner.execute(current.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(runner.history).toHaveLength(1);
  });

  it.each([
    [
      '数字变成错误编码，值和revision都未变',
      (s: SelectionSnapshot): void => {
        s.cellTypes![0]![0] = 'error';
      },
    ],
    [
      '类型读取失败',
      (s: SelectionSnapshot): void => {
        s.cellTypes = undefined;
      },
    ],
    [
      '不能确认全部行',
      (s: SelectionSnapshot): void => {
        s.rowScope = undefined;
      },
    ],
    [
      '单元格值变化',
      (s: SelectionSnapshot): void => {
        s.values![0]![0] = 999;
      },
    ],
    [
      '范围变化',
      (s: SelectionSnapshot): void => {
        s.targetId = 'costs!C2:C5';
      },
    ],
    [
      '文档变化',
      (s: SelectionSnapshot): void => {
        s.documentId = 'another-workbook';
      },
    ],
  ] as const)('%s时不展示过期结果', async (_label, mutate) => {
    const host = new MockAdapter('sheet');
    const state = await host.readSelection({ includeCellTypes: true });
    vi.spyOn(host, 'readSelection').mockImplementation(async () => structuredClone(state));
    const apply = vi.spyOn(host, 'apply');
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    const preview = await runner.prepare('sheet.average');
    mutate(state);

    await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(runner.history).toEqual([]);
    expect(apply).not.toHaveBeenCalled();
  });

  it('现有合计不请求额外类型信息，仍返回450.5', async () => {
    const host = new MockAdapter('sheet');
    const read = vi.spyOn(host, 'readSelection');
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    const result = await runner.execute((await runner.prepare('sheet.sum')).id);
    expect(result.message).toContain('450.5');
    expect(read.mock.calls).toEqual([[undefined], [undefined]]);
  });

  it('演示快照的类型数组与内部状态隔离，普通读取不附带类型', async () => {
    const host = new MockAdapter('sheet');
    const first = await host.readSelection({ includeCellTypes: true });
    first.cellTypes![0]![0] = 'error';
    expect((await host.readSelection({ includeCellTypes: true })).cellTypes?.[0]?.[0]).toBe(
      'number',
    );
    expect(await host.readSelection()).not.toHaveProperty('cellTypes');
  });
});
