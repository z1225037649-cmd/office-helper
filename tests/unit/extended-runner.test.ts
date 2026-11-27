import { describe, expect, it, vi } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import type { FeatureDefinition, SelectionSnapshot } from '../../src/core/contracts';
import { validateExtendedPlan } from '../../src/core/extended-actions';
import { FeatureRegistry } from '../../src/core/registry';
import { TaskRunner } from '../../src/core/runner';
import { features } from '../../src/features';
import { LocalIntentProvider } from '../../src/services/intent';

const setup = (kind: 'text' | 'sheet' = 'sheet') => {
  const host = new MockAdapter(kind);
  return { host, runner: new TaskRunner(host, new FeatureRegistry(features)) };
};
const notice = {
  title: '周末活动通知',
  time: '2026年11月28日 09:00',
  location: '活动室',
  body: '请自带水杯。\n欢迎参加。',
  contact: '张老师 12345',
};

describe('扩展任务执行与安全边界', () => {
  it.each(['sheet.autofit', 'sheet.wrap', 'sheet.header', 'sheet.print'])(
    '%s 先预览、确认后只改布局并可恢复',
    async (id) => {
      const { host, runner } = setup();
      const options = { includeSheetLayout: true, includePrintSettings: true };
      const before = await host.readSelection(options);
      const task = await runner.prepare(id);
      expect(await host.readSelection(options)).toEqual(before);
      expect((await runner.execute(task.id)).canUndo).toBe(true);
      expect((await host.readSelection()).values).toEqual(before.values);
      await runner.undo();
      const after = await host.readSelection(options);
      expect(after.sheetLayout).toEqual(before.sheetLayout);
      expect(after.printSettings).toEqual(before.printSettings);
      expect(after.values).toEqual(before.values);
    },
  );

  it('仅原始列宽变化也会使预览失效', async () => {
    const { host, runner } = setup();
    const task = await runner.prepare('sheet.wrap');
    const read = host.readSelection.bind(host);
    vi.spyOn(host, 'readSelection').mockImplementation(async (options) => {
      const snapshot = await read(options);
      if (snapshot.sheetLayout) snapshot.sheetLayout.columnWidths[0] = 777;
      return snapshot;
    });
    await expect(runner.execute(task.id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect((await read({ includeSheetLayout: true })).sheetLayout?.wrapText).toEqual([
      [false],
      [false],
      [false],
      [false],
    ]);
  });

  it('表单打开后选区改变，不将原列选择应用到新范围', async () => {
    const { host, runner } = setup();
    const before = await host.readSelection();
    const changed = { ...before, targetId: 'different' };
    vi.spyOn(host, 'readSelection').mockResolvedValue(changed);
    await expect(
      runner.prepare('document.saveCopy', { fileName: '副本.xlsx' }, before),
    ).rejects.toMatchObject({ code: 'STALE_SELECTION' });
  });

  it('打印接口不可用的只读指引显示guidance，不能记为完成设置', async () => {
    const { host, runner } = setup();
    const read = host.readSelection.bind(host);
    vi.spyOn(host, 'readSelection').mockImplementation(async (options) => {
      const snapshot = await read(options);
      if (options?.includePrintSettings)
        snapshot.printSettings = {
          available: false,
          reason: '无接口',
          printArea: '',
          orientation: '',
          zoom: false,
          fitWide: 0,
          fitTall: 0,
        };
      return snapshot;
    });
    const result = await runner.execute((await runner.prepare('sheet.print')).id);
    expect(result.outcome).toBe('guidance');
    expect(result.canUndo).toBe(false);
    expect(runner.history).toHaveLength(0);
  });

  it('通知生成独立文档，保留原稿和之前的恢复凭据', async () => {
    const { host, runner } = setup('text');
    await runner.execute((await runner.prepare('text.enlarge')).id);
    const before = await host.readSelection();
    const task = await runner.prepare('document.notice', { notice });
    expect(host.createdDocuments).toHaveLength(0);
    expect(task.confirmLabel).toBeTruthy();
    await runner.execute(task.id);
    expect(await host.readSelection()).toEqual(before);
    expect(host.createdDocuments[0]?.text).toContain('请自带水杯。\n欢迎参加。');
    expect(runner.canUndo).toBe(true);
    await runner.undo();
    expect((await host.readSelection()).fontSize).toBe(18);
  });

  it.each(['guidance', 'cancelled', 'download'] as const)(
    '导出结果 %s 不记为已保存，不清除原稿恢复',
    async (outcome) => {
      const { host, runner } = setup('text');
      await runner.execute((await runner.prepare('text.enlarge')).id);
      const historyCount = runner.history.length;
      vi.spyOn(host, 'apply').mockResolvedValueOnce({
        token: 'external',
        canUndo: false,
        outcome,
        message: '请核对原生窗口。',
      });
      const result = await runner.execute(
        (await runner.prepare('document.exportPdf', { fileName: '通知.pdf' })).id,
      );
      expect(result.outcome).toBe(outcome);
      expect(runner.history).toHaveLength(historyCount);
      expect(runner.canUndo).toBe(true);
    },
  );

  it('检查结果导航只能使用本次报告内的标识，原始数据改变后拒绝定位', async () => {
    const { host } = setup();
    const feature: FeatureDefinition = {
      id: 'report',
      title: '报告',
      description: '',
      icon: 'inspect',
      kinds: ['sheet'],
      requires: ['selection.read'],
      phrases: [],
      prepare: () => ({
        summary: '检查',
        plan: {
          type: 'report',
          message: '需核对',
          findings: [{ id: 'row-2', label: '第2行', description: '疑似重复', row: 1, column: 0 }],
        },
      }),
    };
    const runner = new TaskRunner(host, new FeatureRegistry([feature]));
    const result = await runner.execute((await runner.prepare(feature.id)).id);
    expect(result.findings).toHaveLength(1);
    await runner.locate('row-2');
    expect(host.locatedCell).toEqual({ row: 1, column: 0 });
    await expect(runner.locate('arbitrary-row')).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    const read = host.readSelection.bind(host);
    vi.spyOn(host, 'readSelection').mockImplementation(async (options) => ({
      ...(await read(options)),
      documentId: 'other-file',
    }));
    await expect(runner.locate('row-2')).rejects.toMatchObject({ code: 'STALE_SELECTION' });
  });

  it('拒绝越过选区的报告位置', async () => {
    const feature: FeatureDefinition = {
      id: 'bad',
      title: '报告',
      description: '',
      icon: 'inspect',
      kinds: ['sheet'],
      requires: [],
      phrases: [],
      prepare: () => ({
        summary: '检查',
        plan: {
          type: 'report',
          message: '检查',
          findings: [{ id: 'bad', label: '', description: '', row: 999, column: 0 }],
        },
      }),
    };
    const runner = new TaskRunner(new MockAdapter('sheet'), new FeatureRegistry([feature]));
    await expect(runner.prepare('bad')).rejects.toMatchObject({ code: 'INVALID_PLAN' });
  });

  it('扩展动作也不能漏报能力绕过宿主限制', async () => {
    const feature: FeatureDefinition = {
      id: 'bad',
      title: '修改',
      description: '',
      icon: 'inspect',
      kinds: ['text'],
      requires: [],
      phrases: [],
      prepare: () => ({ summary: '修改', plan: { type: 'sheet.wrap' } }),
    };
    const runner = new TaskRunner(new MockAdapter('text'), new FeatureRegistry([feature]));
    await expect(runner.prepare('bad')).rejects.toMatchObject({ code: 'UNAVAILABLE' });
  });

  it('结果按钮仅能调用最近输出允许的宿主标识', async () => {
    const host = Object.assign(new MockAdapter('text'), {
      openResult: vi.fn(async (_id: string) => undefined),
    });
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    vi.spyOn(host, 'apply').mockResolvedValueOnce({
      token: 'file',
      canUndo: false,
      outcome: 'download',
      message: '已交给浏览器。',
      actions: [{ id: 'opaque-result', label: '打开结果' }],
    });
    const result = await runner.execute(
      (await runner.prepare('document.exportPdf', { fileName: '副本.pdf' })).id,
    );
    expect(result.actions).toEqual([{ id: 'opaque-result', label: '打开结果' }]);
    await expect(runner.openResult('C:\\other.pdf')).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
    expect(host.openResult).not.toHaveBeenCalled();
    await runner.openResult('opaque-result');
    expect(host.openResult).toHaveBeenCalledWith('opaque-result');
    await runner.prepare('selection.inspect');
    await expect(runner.openResult('opaque-result')).rejects.toMatchObject({
      code: 'EXPIRED_TASK',
    });
  });
});

describe('文件名及模板参数校验', () => {
  it.each(['../原稿.pdf', 'C:\\原稿.pdf', 'CON.pdf', '副本?.pdf', '副本.pdf ', '副本.docx'])(
    '拒绝不安全 PDF 文件名 %s',
    (fileName) => {
      expect(() => validateExtendedPlan({ type: 'document.exportPdf', fileName })).toThrow();
    },
  );
  it('拒绝空白必填项与正文控制字符', () => {
    expect(() =>
      validateExtendedPlan({ type: 'document.notice', fields: { ...notice, contact: ' ' } }),
    ).toThrow();
    expect(() =>
      validateExtendedPlan({ type: 'document.notice', fields: { ...notice, body: '\u0000' } }),
    ).toThrow();
  });
  it('普通读取不携带额外布局数据', async () => {
    const snapshot: SelectionSnapshot = await new MockAdapter('sheet').readSelection();
    expect(snapshot.sheetLayout).toBeUndefined();
    expect(snapshot.printSettings).toBeUndefined();
  });
  it('包含连词的单项检查短句能打开参数流程', async () => {
    const provider = new LocalIntentProvider();
    expect(await provider.resolve('找出重复和漏填', features)).toMatchObject({
      type: 'match',
      featureId: 'sheet.checkRecords',
    });
  });
});
