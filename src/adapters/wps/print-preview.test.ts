import { describe, expect, it, vi } from 'vitest';
import type {
  CellValue,
  ChangeReceipt,
  SelectionReadOptions,
  SelectionSnapshot,
} from '../../core/contracts';
import { FeatureRegistry } from '../../core/registry';
import { TaskRunner } from '../../core/runner';
import { printCheck } from '../../features/print-check/feature';
import { WpsAdapter } from './adapter';
import type { WpsApplication, WpsPageSetup, WpsSelection, WpsWorkbook } from './bridge';
import { withWpsPrintPreview, type WpsPrintPreviewWorksheet } from './print-preview';

function fixture() {
  const values: CellValue[][] = [
    ['项目', '金额'],
    ['甲', 25],
  ];
  const book: WpsWorkbook = {
    Name: '打印测试.xlsx',
    FullName: 'C:\\fixture\\打印测试.xlsx',
    ReadOnly: false,
  };
  const page: WpsPageSetup = {
    PrintArea: '$A$1:$G$40',
    Orientation: 1,
    Zoom: 90,
    FitToPagesWide: 2,
    FitToPagesTall: 3,
  };
  const preview = vi.fn(function (this: WpsPrintPreviewWorksheet, _enableChanges?: boolean) {
    expect(this).toBe(sheet);
  });
  const forbidden = vi.fn(() => {
    throw new Error('A printing, activation, shell or command API must never be used');
  });
  const sheet: WpsPrintPreviewWorksheet = {
    Name: '费用',
    Parent: book,
    ProtectContents: false,
    PageSetup: page,
    PrintPreview: preview,
    Activate: forbidden,
  };
  Object.defineProperty(sheet, 'PrintOut', { get: forbidden });
  const range: WpsSelection = {
    Row: 1,
    Column: 1,
    Rows: { Count: 2 },
    Columns: { Count: 2 },
    Areas: { Count: 1 },
    Address: '$A$1:$B$2',
    Worksheet: sheet,
    get Value2() {
      return structuredClone(values);
    },
    Cells: {
      Item: (row, column) => ({
        Value2: values[row - 1]![column - 1],
        ColumnWidth: 12,
        WrapText: false,
        Font: { Bold: false },
        MergeCells: false,
      }),
    },
  };
  sheet.Range = () => range;
  const app: WpsApplication = {
    ActiveWorkbook: book,
    ActiveSheet: sheet,
    Selection: range,
    ActiveWindow: { Hwnd: 123 },
    Enum: { xlLandscape: 2, xlPortrait: 1 },
    OAAssist: { ShellExecute: forbidden },
  };
  for (const name of ['PrintOut', 'CommandBars', 'Run', 'SendKeys'])
    Object.defineProperty(app, name, { get: forbidden });
  const native = new WpsAdapter('sheet', () => app);
  const host = withWpsPrintPreview(native, () => app);
  const runner = new TaskRunner(host, new FeatureRegistry([printCheck]));
  const execute = async () => runner.execute((await runner.prepare('sheet.print')).id);
  const direct = async () => {
    const expected = await host.readSelection({ includePrintSettings: true });
    return host.apply({ type: 'sheet.print' }, expected);
  };
  return {
    values,
    book,
    page,
    sheet,
    range,
    app,
    preview,
    forbidden,
    native,
    host,
    runner,
    execute,
    direct,
  };
}

function actionId(result: { actions?: ChangeReceipt['actions'] }): string {
  expect(result.actions).toEqual([
    { id: expect.stringMatching(/^wps-print-preview:/u), label: '打开原生打印预览（不打印）' },
  ]);
  return result.actions![0]!.id;
}

describe('WPS 安全原生打印预览入口', () => {
  it('真实适配器与 runner：先确认设置，独立点击结果才调用 PrintPreview(false)，并保留恢复', async () => {
    const f = fixture();
    const original = structuredClone(f.page);
    const task = await f.runner.prepare('sheet.print');
    expect(f.page).toEqual(original);
    expect(f.preview).not.toHaveBeenCalled();
    const result = await f.runner.execute(task.id);
    const id = actionId(result);
    expect(result).toMatchObject({ canUndo: true, outcome: 'completed' });
    expect(f.preview).not.toHaveBeenCalled();
    expect(f.page).toMatchObject({
      PrintArea: '$A$1:$B$2',
      Orientation: 2,
      Zoom: false,
      FitToPagesWide: 1,
      // WPS stores automatic height as false; the public snapshot normalizes it to zero.
      FitToPagesTall: false,
    });
    const configured = structuredClone(f.page);
    await f.runner.openResult(id);
    expect(f.preview).toHaveBeenCalledExactlyOnceWith(false);
    expect(f.page).toEqual(configured);
    expect(f.values).toEqual([
      ['项目', '金额'],
      ['甲', 25],
    ]);
    expect(f.runner.canUndo).toBe(true);
    expect(f.runner.history).toHaveLength(1);
    await f.runner.undo();
    expect(f.page).toEqual(original);
    await expect(f.runner.openResult(id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(f.forbidden).not.toHaveBeenCalled();
  });

  it('取消确认不设置参数、不创建结果入口，也不进入预览', async () => {
    const f = fixture();
    const original = structuredClone(f.page);
    const task = await f.runner.prepare('sheet.print');
    f.runner.cancel(task.id);
    await expect(f.runner.execute(task.id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(f.page).toEqual(original);
    expect(f.preview).not.toHaveBeenCalled();
  });

  it.each(['missing', 'throwing getter'] as const)(
    '预览方法%s时保留已完成设置及恢复凭据，只给原生指引',
    async (mode) => {
      const f = fixture();
      if (mode === 'missing') delete f.sheet.PrintPreview;
      else
        Object.defineProperty(f.sheet, 'PrintPreview', {
          get: () => {
            throw new Error('unsupported bridge');
          },
        });
      const original = structuredClone(f.page);
      const result = await f.execute();
      expect(result.actions).toBeUndefined();
      expect(result).toMatchObject({ canUndo: true, outcome: 'completed' });
      expect(result.message).toContain('未提供自动预览按钮');
      expect(result.message).toContain('文件 → 打印');
      expect(result.message).toContain('Ctrl+P');
      await f.runner.undo();
      expect(f.page).toEqual(original);
      expect(f.preview).not.toHaveBeenCalled();
      expect(f.forbidden).not.toHaveBeenCalled();
    },
  );

  it('打印设置接口缺失时只读 report，不因有预览方法就绕过原任务生成按钮', async () => {
    const f = fixture();
    delete f.sheet.PageSetup;
    const result = await f.execute();
    expect(result).toMatchObject({ outcome: 'guidance', canUndo: false });
    expect(result.actions).toBeUndefined();
    expect(f.runner.history).toHaveLength(0);
    expect(f.preview).not.toHaveBeenCalled();
  });

  it.each([
    ['工作簿', (f: ReturnType<typeof fixture>) => (f.app.ActiveWorkbook = { ...f.book })],
    ['工作表', (f: ReturnType<typeof fixture>) => (f.app.ActiveSheet = { ...f.sheet })],
    ['窗口', (f: ReturnType<typeof fixture>) => (f.app.ActiveWindow = { Hwnd: 999 })],
    ['选区', (f: ReturnType<typeof fixture>) => (f.range.Address = '$C$3:$D$4')],
    ['内容', (f: ReturnType<typeof fixture>) => (f.values[1]![1] = 888)],
    ['打印区域', (f: ReturnType<typeof fixture>) => (f.page.PrintArea = '$A$1:$G$40')],
    ['方向', (f: ReturnType<typeof fixture>) => (f.page.Orientation = 1)],
    ['缩放', (f: ReturnType<typeof fixture>) => (f.page.Zoom = 75)],
    ['页宽', (f: ReturnType<typeof fixture>) => (f.page.FitToPagesWide = 2)],
    ['页高', (f: ReturnType<typeof fixture>) => (f.page.FitToPagesTall = 2)],
  ])('结果生成后%s改变则拒绝进入预览，不激活、重设或打印', async (_name, change) => {
    const f = fixture();
    const id = actionId(await f.execute());
    change(f);
    const before = structuredClone(f.page);
    await expect(f.runner.openResult(id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(f.page).toEqual(before);
    expect(f.preview).not.toHaveBeenCalled();
    expect(f.forbidden).not.toHaveBeenCalled();
  });

  it('结果复核跨越异步等待，期间工作表改变仍会在原生调用前拒绝', async () => {
    const f = fixture();
    const id = actionId(await f.execute());
    const read = f.native.readSelection.bind(f.native);
    vi.spyOn(f.native, 'readSelection').mockImplementation(async (options) => {
      const snapshot = await read(options);
      f.app.ActiveSheet = { ...f.sheet };
      return snapshot;
    });
    await expect(f.host.openResult!(id)).rejects.toMatchObject({ code: 'STALE_SELECTION' });
    expect(f.preview).not.toHaveBeenCalled();
  });

  it('原生预览方法在点击前消失时给指引，绝不转用 PrintOut 或命令执行', async () => {
    const f = fixture();
    const id = actionId(await f.execute());
    delete f.sheet.PrintPreview;
    await expect(f.runner.openResult(id)).rejects.toMatchObject({
      code: 'UNAVAILABLE',
      message: expect.stringContaining('文件 → 打印'),
    });
    expect(f.preview).not.toHaveBeenCalled();
    expect(f.forbidden).not.toHaveBeenCalled();
  });

  it('预览方法 getter 在点击时抛错，仍提供明确原生指引并作废该入口', async () => {
    const f = fixture();
    const id = actionId(await f.execute());
    Object.defineProperty(f.sheet, 'PrintPreview', {
      get: () => {
        throw new Error('bridge disconnected');
      },
    });
    await expect(f.runner.openResult(id)).rejects.toMatchObject({
      code: 'UNAVAILABLE',
      message: expect.stringContaining('文件 → 打印'),
    });
    await expect(f.runner.openResult(id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(f.preview).not.toHaveBeenCalled();
    expect(f.forbidden).not.toHaveBeenCalled();
  });

  it('WPS 每次提供新的原生包装对象而无法确认身份时，保守降级手动指引', async () => {
    const f = fixture();
    Object.defineProperty(f.app, 'ActiveSheet', { get: () => ({ ...f.sheet }) });
    const result = await f.execute();
    expect(result).toMatchObject({ outcome: 'completed', canUndo: true });
    expect(result.actions).toBeUndefined();
    expect(result.message).toContain('未提供自动预览按钮');
    expect(f.preview).not.toHaveBeenCalled();
  });

  it('异步复核期间 dispose 后，不能用复核前的有效结果调用原生预览', async () => {
    const f = fixture();
    const id = actionId(await f.execute());
    const read = f.native.readSelection.bind(f.native);
    vi.spyOn(f.native, 'readSelection').mockImplementation(async (options) => {
      const snapshot = await read(options);
      f.host.dispose();
      return snapshot;
    });
    await expect(f.host.openResult!(id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(f.preview).not.toHaveBeenCalled();
  });

  it.each(['throw', 'reject'] as const)(
    '原生方法%s后不谎报已打开、不重试、保留恢复',
    async (mode) => {
      const f = fixture();
      f.sheet.PrintPreview = vi.fn(() => {
        if (mode === 'throw') throw new Error('native failure');
        return Promise.reject(new Error('native failure'));
      });
      const id = actionId(await f.execute());
      await expect(f.runner.openResult(id)).rejects.toMatchObject({
        code: 'UNAVAILABLE',
        message: expect.stringContaining('未能确认原生打印预览已打开'),
      });
      await expect(f.runner.openResult(id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
      expect(f.sheet.PrintPreview).toHaveBeenCalledExactlyOnceWith(false);
      expect(f.runner.canUndo).toBe(true);
      await f.runner.undo();
      expect(f.page.Zoom).toBe(90);
      expect(f.forbidden).not.toHaveBeenCalled();
    },
  );

  it('不接收伪造或旧预览 ID，非打印结果仍精确委托原适配器', async () => {
    const f = fixture();
    const delegated = vi.spyOn(f.native, 'openResult').mockResolvedValue();
    const first = actionId(await f.direct());
    const second = actionId(await f.direct());
    for (const id of [first, 'wps-print-preview:made-up'])
      await expect(f.host.openResult!(id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
    expect(delegated).not.toHaveBeenCalled();
    await f.host.openResult!('wps-result:exported-pdf');
    expect(delegated).toHaveBeenCalledExactlyOnceWith('wps-result:exported-pdf');
    await f.host.openResult!(second);
    expect(f.preview).toHaveBeenCalledExactlyOnceWith(false);
  });

  it.each(['undo', 'dispose', 'apply failure'] as const)(
    '%s立即使旧入口失效',
    async (operation) => {
      const f = fixture();
      const receipt = await f.direct();
      const id = actionId(receipt);
      if (operation === 'undo') await f.host.undo(receipt);
      else if (operation === 'dispose') f.host.dispose();
      else {
        vi.spyOn(f.native, 'apply').mockRejectedValue(new Error('failed'));
        await expect(
          f.host.apply({ type: 'sheet.print' }, await f.host.readSelection()),
        ).rejects.toThrow('failed');
      }
      await expect(f.host.openResult!(id)).rejects.toMatchObject({ code: 'EXPIRED_TASK' });
      expect(f.preview).not.toHaveBeenCalled();
    },
  );

  it('设置后的额外回读失败不丢失真实结果 token、恢复能力及已有结果按钮', async () => {
    const f = fixture();
    const expected = await f.host.readSelection({ includePrintSettings: true });
    const receipt = await f.native.apply({ type: 'sheet.print' }, expected);
    receipt.actions = [{ id: 'existing', label: '已有结果' }];
    vi.spyOn(f.native, 'apply').mockResolvedValue(receipt);
    vi.spyOn(f.native, 'readSelection').mockRejectedValue(new Error('read failed'));
    const result = await f.host.apply({ type: 'sheet.print' }, expected);
    expect(result).toMatchObject({
      token: receipt.token,
      canUndo: true,
      outcome: 'completed',
      actions: receipt.actions,
    });
    expect(result.message).toContain('未提供自动预览按钮');
    expect(f.preview).not.toHaveBeenCalled();
  });

  it.each(['guidance', 'cancelled', 'download'] as const)(
    '%s结果不注册打印入口',
    async (outcome) => {
      const f = fixture();
      const receipt: ChangeReceipt = {
        token: 'untouched',
        canUndo: false,
        message: outcome,
        outcome,
      };
      vi.spyOn(f.native, 'apply').mockResolvedValue(receipt);
      expect(await f.direct()).toBe(receipt);
      expect(f.preview).not.toHaveBeenCalled();
    },
  );

  it('只读取原快照要求的附加数据，并在点击时重复同样的读取范围', async () => {
    const f = fixture();
    const read = vi.spyOn(f.native, 'readSelection');
    const expected = await f.host.readSelection({
      includePrintSettings: true,
      includeCellTypes: false,
    });
    const result = await f.host.apply({ type: 'sheet.print' }, expected);
    const options: SelectionReadOptions = { includePrintSettings: true, includeSheetLayout: true };
    expect(read).toHaveBeenLastCalledWith(options);
    await f.host.openResult!(actionId(result));
    expect(read).toHaveBeenLastCalledWith(options);
  });

  it('已有非打印功能原样委托，不改变其凭据、结果路由或历史', async () => {
    const f = fixture();
    const receipt: ChangeReceipt = {
      token: 'export-result',
      canUndo: false,
      message: 'exported',
      outcome: 'completed',
      actions: [{ id: 'wps-result:pdf', label: 'PDF' }],
    };
    vi.spyOn(f.native, 'apply').mockResolvedValue(receipt);
    const result = await f.host.apply(
      { type: 'document.exportPdf', fileName: '副本.pdf' },
      await f.host.readSelection(),
    );
    expect(result).toBe(receipt);
    expect(f.preview).not.toHaveBeenCalled();
  });

  it('Office、演示与文字宿主保持原对象，不访问 WPS 全局接口', () => {
    const f = fixture();
    const application = vi.fn(() => {
      throw new Error('must not access WPS');
    });
    for (const info of [
      { ...f.native.info, id: 'office' as const },
      { ...f.native.info, id: 'mock' as const },
      { ...f.native.info, kind: 'text' as const },
    ]) {
      const host = Object.assign(Object.create(f.native), { info });
      expect(withWpsPrintPreview(host, application)).toBe(host);
    }
    expect(application).not.toHaveBeenCalled();
  });

  it('同一入口并发点击只请求一次；原生无返回值也不声称已确认窗口或页数', async () => {
    const f = fixture();
    let finish!: () => void;
    f.sheet.PrintPreview = vi.fn(() => new Promise<void>((resolve) => (finish = resolve)));
    const id = actionId(await f.direct());
    const pending = f.host.openResult!(id);
    await vi.waitFor(() => expect(f.sheet.PrintPreview).toHaveBeenCalledOnce());
    await expect(f.host.openResult!(id)).rejects.toMatchObject({ code: 'BUSY' });
    finish();
    expect(await pending).toBeUndefined();
    expect(f.sheet.PrintPreview).toHaveBeenCalledExactlyOnceWith(false);
    expect(f.forbidden).not.toHaveBeenCalled();
  });

  it.each(['PrintArea', 'Orientation', 'Zoom', 'FitToPagesWide', 'FitToPagesTall'] as const)(
    '成功凭据若无法回读确认目标参数%s，不冒险创建入口',
    async (property) => {
      const f = fixture();
      const expected = await f.host.readSelection({ includePrintSettings: true });
      const apply = f.native.apply.bind(f.native);
      vi.spyOn(f.native, 'apply').mockImplementation(async (plan, source) => {
        const result = await apply(plan, source);
        const original = {
          PrintArea: '$A$1:$G$40',
          Orientation: 1,
          Zoom: 90,
          FitToPagesWide: 2,
          FitToPagesTall: 3,
        };
        Object.assign(f.page, { [property]: original[property] });
        return result;
      });
      const result = await f.host.apply({ type: 'sheet.print' }, expected);
      expect(result.actions).toBeUndefined();
      expect(result.canUndo).toBe(true);
      expect(result.message).toContain('未提供自动预览按钮');
      expect(f.preview).not.toHaveBeenCalled();
    },
  );
});

describe('打印功能不把参数或尺寸当作实际分页', () => {
  it('一页宽、自动页高、原生内容宽度只描述设置，不计算页数或缩放比例', async () => {
    const f = fixture();
    const snapshot: SelectionSnapshot = await f.host.readSelection({ includePrintSettings: true });
    snapshot.printSettings!.contentWidthPoints = 12345;
    const before = structuredClone(snapshot);
    const result = printCheck.prepare(snapshot);
    expect(result.summary).toContain('这里的页数参数是设置值');
    expect(result.summary).toContain('实际页数需在原生打印预览中核对');
    expect(result.summary).toContain('确认仅设置打印参数');
    expect(result.summary).toContain('若结果提供');
    expect(result.summary).toContain('文字过小');
    expect(result.summary).not.toMatch(/(?:实际|预计|总计|总共|共)\s*\d+\s*页/u);
    expect(result.summary).not.toContain('12345');
    expect(snapshot).toEqual(before);
    expect(f.preview).not.toHaveBeenCalled();
  });
});
