import { describe, expect, it } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import { createAdapter } from '../../src/adapters';
import { FeatureRegistry } from '../../src/core/registry';
import { TaskRunner } from '../../src/core/runner';
import { features } from '../../src/features';
import { sumSelection } from '../../src/features/sum-selection/feature';
import { LocalIntentProvider } from '../../src/services/intent';

describe('功能注册与指令解析', () => {
  it('重复 ID 在启动时暴露，而不是覆盖旧功能', () => {
    expect(() => new FeatureRegistry([features[0]!, features[0]!])).toThrow('重复');
  });
  it('关闭功能后入口和执行权限同时关闭', () => {
    const registry = new FeatureRegistry(features, { 'text.enlarge': false });
    expect(registry.list().some((feature) => feature.id === 'text.enlarge')).toBe(false);
    expect(() => registry.get('text.enlarge')).toThrow();
  });
  it('普通表达映射到注册功能，多任务需要用户选择', async () => {
    const provider = new LocalIntentProvider();
    expect(await provider.resolve('请帮我把字放大一点。', features)).toEqual({
      type: 'match',
      featureId: 'text.enlarge',
    });
    expect(await provider.resolve('把字放大到30', features)).toEqual({
      type: 'match',
      featureId: 'text.enlarge',
      parameters: { fontSize: 30 },
    });
    expect(await provider.resolve('把字放大到100磅', features)).toMatchObject({
      type: 'clarify',
    });
    expect(await provider.resolve('把字放大到', features)).toMatchObject({
      type: 'clarify',
    });
    expect(await provider.resolve('把字放大一号', features)).toMatchObject({
      type: 'clarify',
    });
    expect(await provider.resolve('把字放大到30厘米', features)).toMatchObject({
      type: 'clarify',
    });
    expect(await provider.resolve('把字放大，然后求和', features)).toMatchObject({
      type: 'choices',
    });
    expect(await provider.resolve('删除所有文件', features)).toEqual({ type: 'unknown' });
    expect(await provider.resolve('把字放大', [])).toEqual({ type: 'unknown' });
  });
  it('明确的宿主入口失败不能伪装成演示成功', async () => {
    await expect(createAdapter('?host=unexpected')).rejects.toMatchObject({
      code: 'HOST_CONNECTION',
    });
    await expect(createAdapter('?host=wps')).rejects.toMatchObject({ code: 'HOST_CONNECTION' });
  });
});

describe('计算合计', () => {
  it('使用选区数据得到合计，且不会改写工作表', async () => {
    const host = new MockAdapter('sheet');
    const before = await host.readSelection();
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    const result = await runner.execute((await runner.prepare('sheet.sum')).id);
    expect(result.message).toContain('450.5');
    expect(result.canUndo).toBe(false);
    expect(await host.readSelection()).toEqual(before);
  });
  it('不把字符串、空白、布尔值和错误值当数字', async () => {
    const selection = await new MockAdapter('sheet').readSelection();
    selection.values = [[0.1, 0.2, 0, '500', '#VALUE!', null, true]];
    const { plan } = sumSelection.prepare(selection);
    expect(plan).toMatchObject({ type: 'report' });
    if (plan.type === 'report') {
      expect(plan.message).toContain('合计 0.3');
      expect(plan.message).toContain('4 个');
    }
  });
  it('空选区不会误报合计为零', async () => {
    const selection = await new MockAdapter('sheet').readSelection();
    selection.values = [['标题', null]];
    expect(() => sumSelection.prepare(selection)).toThrow('没有数字');
  });
});
