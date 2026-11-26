import { describe, expect, it, vi } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import { WpsAdapter } from '../../src/adapters/wps/adapter';
import { FeatureRegistry } from '../../src/core/registry';
import { TaskRunner } from '../../src/core/runner';
import { features } from '../../src/features';
import { validateTextPlan } from '../../src/core/text-format';
import type { ActionPlan } from '../../src/core/contracts';

const setup = () => {
  const host = new MockAdapter();
  return { host, runner: new TaskRunner(host, new FeatureRegistry(features)) };
};

describe('文字格式的预览、确认和恢复', () => {
  it.each([
    ['text.bold', {}, 'bold', true],
    ['paragraph.center', {}, 'alignment', 'Centered'],
    ['paragraph.spacing', { lineSpacingMultiple: 1.5 }, 'lineSpacing', 18],
    ['paragraph.indent', { firstLineIndentPoints: 24 }, 'firstLineIndent', 24],
  ] as const)('%s 预览不写入，确认后只改指定格式并能恢复', async (id, parameters, key, value) => {
    const { host, runner } = setup();
    const original = await host.readSelection();
    const preview = await runner.prepare(id, parameters);
    expect(await host.readSelection()).toEqual(original);
    expect(preview.summary).toContain(id === 'text.bold' ? '常规 → 加粗' : '完整段落');
    await runner.execute(preview.id);
    const after = await host.readSelection();
    expect(key === 'bold' ? after.bold : after.paragraphs![0]![key]).toBe(value);
    expect(after.text).toBe(original.text);
    expect(after.fontSize).toBe(original.fontSize);
    await runner.undo();
    expect({ ...(await host.readSelection()), revision: original.revision }).toEqual(original);
  });

  it.each(['paragraph.spacing', 'paragraph.indent'])(
    '%s 缺少明确参数时要求说明，不生成默认修改',
    async (id) => {
      const { host, runner } = setup();
      const before = await host.readSelection();
      await expect(runner.prepare(id)).rejects.toMatchObject({ code: 'INVALID_PLAN' });
      expect(await host.readSelection()).toEqual(before);
    },
  );

  it('同段未选中文字或其他格式发生变化后拒绝旧预览', async () => {
    const { host, runner } = setup();
    const originalRead = host.readSelection.bind(host);
    const preview = await runner.prepare('paragraph.center');
    vi.spyOn(host, 'readSelection').mockImplementation(async () => {
      const next = await originalRead();
      next.paragraphs![0]!.text += '后续编辑';
      return next;
    });
    await expect(runner.execute(preview.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
  });

  it('未知原始行距、混合粗细和编号缩进被阻止', async () => {
    const snapshot = await new MockAdapter().readSelection();
    snapshot.bold = undefined;
    snapshot.paragraphs![0]!.spacingMode = 'unsupported';
    snapshot.paragraphs![0]!.isListItem = true;
    for (const [id, parameters] of [
      ['text.bold', {}],
      ['paragraph.spacing', { lineSpacingMultiple: 1.5 }],
      ['paragraph.indent', { firstLineIndentPoints: 24 }],
    ] as const)
      expect(() => features.find((f) => f.id === id)!.prepare(snapshot, parameters)).toThrow();
  });

  it('WPS 文字开放全部文字格式，表格不提供文字能力', () => {
    const registry = new FeatureRegistry(features);
    const writer = new WpsAdapter('text');
    expect(registry.unavailableReason(registry.get('text.bold'), writer.info)).toBeUndefined();
    expect(
      registry.unavailableReason(registry.get('paragraph.center'), writer.info),
    ).toBeUndefined();
    for (const kind of ['text', 'sheet'] as const) {
      const host = new WpsAdapter(kind);
      for (const id of ['paragraph.spacing', 'paragraph.indent'])
        expect(!!registry.unavailableReason(registry.get(id), host.info)).toBe(kind === 'sheet');
    }
    const sheet = new WpsAdapter('sheet');
    expect(registry.unavailableReason(registry.get('paragraph.center'), sheet.info)).toBeTruthy();
    expect(registry.unavailableReason(registry.get('text.bold'), sheet.info)).toBeTruthy();
  });

  it.each([
    { type: 'text.bold', bold: 'true' },
    { type: 'paragraph.spacing', multiple: NaN },
    { type: 'paragraph.spacing', multiple: 0.5 },
    { type: 'paragraph.spacing', multiple: 4 },
    { type: 'paragraph.spacing', multiple: 1.234 },
    { type: 'paragraph.indent', points: -1 },
    { type: 'paragraph.indent', points: Infinity },
    { type: 'paragraph.indent', points: 145 },
    { type: 'paragraph.indent', points: 24.024 },
    { type: 'text.resize', size: 30.3 },
    { type: 'delete' },
  ])('执行器拒绝不合法格式参数 %j', (plan) => {
    expect(() => validateTextPlan(plan as ActionPlan)).toThrow();
  });
});
