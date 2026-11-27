import { afterAll, describe, expect, it, vi } from 'vitest';
import { mkdtemp, mkdir, readFile, writeFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname, basename, resolve } from 'node:path';
import JSZip from 'jszip';
import { runInNewContext } from 'node:vm';
import {
  matrices,
  createWpsExtendedEntryMiddleware,
  prepareExtendedCheck,
  rows,
  verifyOutput,
  wpsBootstrap,
} from '../../scripts/prepare-extended-check.mjs';
import {
  assertExtendedConfig,
  extendedRows,
  fixtureText,
  outputName,
} from '../../src/dev/extended-check-spec';

const temporary = [];
async function setup() {
  const root = await mkdtemp(join(tmpdir(), 'office-helper-extended-'));
  temporary.push(root);
  await mkdir(join(root, 'manifests/office'), { recursive: true });
  await mkdir(join(root, '.local'));
  await writeFile(join(root, '.local/native-token'), 'a'.repeat(48));
  await writeFile(
    join(root, 'manifests/office/manifest.xml'),
    await readFile('manifests/office/manifest.xml'),
  );
  return root;
}

// VM doubles only: these tests neither start a server nor connect to a real host.
function bootstrapHarness(script, app, api) {
  const delayed = [];
  const intervals = [];
  const status = {};
  const window = { Application: app, wps: api };
  const clearInterval = vi.fn();
  runInNewContext(script, {
    window,
    document: { getElementById: () => status },
    setTimeout: (callback, delay) => delayed.push({ callback, delay }),
    setInterval: (callback, delay) => intervals.push({ callback, delay }),
    clearInterval,
  });
  return { window, status, delayed, intervals, clearInterval };
}

function fixtureApp(config) {
  const source = {
    Name: config.fileName,
    FullName: config.fixturePath,
    Content: { Text: `${config.text.replaceAll('\n', '\r')}\r` },
  };
  const pane = { Visible: false };
  const app = {
    [config.kind === 'text' ? 'ActiveDocument' : 'ActiveWorkbook']: source,
    ActiveSheet: {
      Name: 'ExtendedData',
      Range: (address) => ({
        Value2:
          address === 'A1:C1'
            ? [['OFFICE_HELPER_EXTENDED_TEST_V1', config.run, config.matrix]]
            : structuredClone(config.values),
      }),
    },
    CreateTaskPane: vi.fn(() => pane),
  };
  return { app, source, pane };
}

function middlewareResponse() {
  return {
    statusCode: 200,
    setHeader: vi.fn(),
    end: vi.fn(),
  };
}
afterAll(async () => {
  for (const path of temporary) {
    if (
      dirname(resolve(path)).toLowerCase() !== resolve(tmpdir()).toLowerCase() ||
      !basename(path).startsWith('office-helper-extended-')
    )
      throw new Error('Unsafe test cleanup');
    await rm(path, { recursive: true, force: true });
  }
});

describe('扩展专项隔离样本准备（文件测试，不代表宿主通过）', () => {
  it('四单元独立、标记逐轮不同，不改 token/已有样本；重复准备永不覆盖', async () => {
    const root = await setup();
    await mkdir(join(root, 'output/native'), { recursive: true });
    const existing = join(root, 'output/native/OfficeHelper-Native-Test.docx');
    await writeFile(existing, 'keep-existing');
    const first = await prepareExtendedCheck(root);
    const second = await prepareExtendedCheck(root);
    expect(first.directory).not.toBe(second.directory);
    expect(first.run).not.toBe(second.run);
    expect(first.configs.map((c) => c.matrix)).toEqual(matrices);
    expect(rows).toEqual(extendedRows);
    for (const config of first.configs) {
      expect(() => assertExtendedConfig(config)).not.toThrow();
      expect(config.text).toBe(fixtureText(first.run, config.matrix));
      expect(await readdir(config.resultDirectory)).toEqual([]);
      expect(config.fixturePath).toContain(config.matrix);
      const zip = await JSZip.loadAsync(await readFile(config.fixturePath), { checkCRC32: true });
      const part = config.kind === 'text' ? 'word/document.xml' : 'xl/worksheets/sheet1.xml';
      const xml = await zip.file(part).async('string');
      expect(xml).toContain(first.run);
      expect(xml).toContain(config.matrix);
      expect(zip.file('_rels/.rels')).not.toBeNull();
      expect(zip.file('[Content_Types].xml')).not.toBeNull();
    }
    expect(await readFile(existing, 'utf8')).toBe('keep-existing');
    expect(await readFile(join(root, '.local/native-token'), 'utf8')).toBe('a'.repeat(48));
    expect(await readdir(join(root, '.local'))).toEqual(['native-token']);
  });

  it('每个 Office 清单唯一 ID、指定宿主、隔离 URL；准备不生成注册/证书文件', async () => {
    const root = await setup();
    const result = await prepareExtendedCheck(root);
    const manifests = await Promise.all(
      ['office-text', 'office-sheet'].map((m) =>
        readFile(join(result.directory, `${m}.xml`), 'utf8'),
      ),
    );
    expect(manifests[0]).toContain('<Host Name="Document" />');
    expect(manifests[0]).not.toContain('<Host Name="Workbook" />');
    expect(manifests[1]).not.toContain('<Host Name="Document" />');
    expect(manifests[1]).toContain('<Host Name="Workbook" />');
    expect(manifests[0].match(/<Id>(.+)<\/Id>/)[1]).not.toBe(
      manifests[1].match(/<Id>(.+)<\/Id>/)[1],
    );
    for (const xml of manifests) {
      expect(xml).toContain(`/__extended/${result.run}/office-`);
      expect(xml).toContain('&amp;token=');
    }
    expect(await readdir(root)).toEqual(expect.arrayContaining(['.local', 'manifests', 'output']));
    expect(await readdir(root)).not.toContain('.office-addin-dev-certs');
    const list = await readFile(join(result.directory, 'CHECKLIST.md'), 'utf8');
    for (const text of [
      'EXPIRED_TASK',
      'STALE_SELECTION',
      'unsupported',
      'ScopeSentinel',
      '正常 UI',
      '不点击覆盖',
    ])
      expect(list).toContain(text);
  });

  it('表格具有偏移矩阵、混合格式、原始打印值、公式、隐藏重复行及第二表哨兵', async () => {
    const result = await prepareExtendedCheck(await setup());
    const config = result.configs.find((c) => c.matrix === 'wps-sheet');
    const zip = await JSZip.loadAsync(await readFile(config.fixturePath));
    const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
    expect(sheet).toContain('sqref="B3:F10"');
    expect(sheet).toContain('<c r="F8"><f>SUM(F4:F5)</f><v>30</v></c>');
    expect(sheet).toContain('<row r="9" hidden="1">');
    expect(sheet).toContain('orientation="portrait" scale="85" fitToWidth="2" fitToHeight="3"');
    expect(sheet).toContain('OFFICE_HELPER_EXTENDED_OUTSIDE');
    expect(sheet).toContain('t="b"><v>0</v>');
    expect(await zip.file('xl/worksheets/sheet2.xml').async('string')).toContain(
      'OFFICE_HELPER_EXTENDED_SECOND_SHEET',
    );
    expect(await zip.file('xl/workbook.xml').async('string')).toContain('ExtendedData!$B$3:$D$6');
  });

  it('缺少/错误 token 时尚未创建任何输出目录', async () => {
    const root = await setup();
    await expect(prepareExtendedCheck(root, 'bad')).rejects.toThrow('token');
    await expect(readdir(join(root, 'output'))).rejects.toMatchObject({ code: 'ENOENT' });
  });

  it('只读结果核对：合法副本仍需人工检查，拒绝缺件、错误标记和被修改的原稿', async () => {
    const result = await prepareExtendedCheck(await setup());
    const config = result.configs[0];
    const path = join(config.resultDirectory, outputName('document.saveCopy', config));
    await writeFile(path, await readFile(config.fixturePath));
    expect(await verifyOutput(config, 'document.saveCopy')).toMatchObject({
      status: 'manual-required',
      path,
    });
    const wrong = new JSZip();
    wrong.file('word/document.xml', '<w:document>wrong</w:document>');
    await writeFile(path, await wrong.generateAsync({ type: 'nodebuffer' }));
    await expect(verifyOutput(config, 'document.saveCopy')).rejects.toThrow('标记');
    await writeFile(config.fixturePath, 'modified');
    await expect(verifyOutput(config, 'document.saveCopy')).rejects.toThrow('原始样本');
  });

  it('PDF 头尾验证不等于全范围通过，不接受任意类型', async () => {
    const result = await prepareExtendedCheck(await setup());
    const config = result.configs[1];
    const path = join(config.resultDirectory, outputName('document.exportPdf', config));
    await writeFile(path, '%PDF-1.7\n%%EOF\n');
    expect((await verifyOutput(config, 'document.exportPdf')).status).toBe('manual-required');
    await writeFile(path, '%PDF-1.7\ntruncated');
    await expect(verifyOutput(config, 'document.exportPdf')).rejects.toThrow('结束');
    await expect(verifyOutput(config, '../../escape')).rejects.toThrow('类型');
  });

  it('独立 WPS 引导只有显式点击才能开窗格，错误文件不创建；不注册或改动文件', async () => {
    const result = await prepareExtendedCheck(await setup());
    const config = result.configs.find((c) => c.matrix === 'wps-text');
    let created = 0;
    const status = {};
    const app = {
      ActiveDocument: {
        Name: config.fileName,
        FullName: config.fixturePath,
        Content: { Text: `${config.text.replaceAll('\n', '\r')}\r` },
      },
      CreateTaskPane(url) {
        expect(url).toBe(config.url);
        created++;
        return {};
      },
    };
    const window = { Application: app };
    runInNewContext(wpsBootstrap(config, config.url), {
      window,
      document: { getElementById: () => status },
    });
    expect(created).toBe(0);
    app.ActiveDocument.Name = 'real-user.docx';
    window.OnOpenExtendedCheck();
    expect(created).toBe(0);
    app.ActiveDocument.Name = config.fileName;
    expect(window.OnOpenExtendedCheck()).toBe(true);
    expect(created).toBe(1);
    window.OnOpenExtendedCheck();
    expect(created).toBe(1);
    expect(await readFile(join(result.directory, 'wps/wps-text/manifest.xml'), 'utf8')).toContain(
      result.run,
    );
    expect(await readFile(join(result.directory, 'WPS-ENTRY.md'), 'utf8')).toContain('blocked');
  });

  it('仅接管已注册 writer/sheet main.js，各自返回本轮专属配置，HTML/ribbon 和其它入口透传', async () => {
    const result = await prepareExtendedCheck(await setup());
    const middleware = createWpsExtendedEntryMiddleware(result.configs);
    for (const [entry, kind] of [
      ['writer', 'text'],
      ['sheet', 'sheet'],
    ]) {
      const response = middlewareResponse();
      const next = vi.fn();
      middleware({ method: 'GET', url: `/wps/${entry}/main.js?cache=1` }, response, next);
      expect(next).not.toHaveBeenCalled();
      expect(response.setHeader).toHaveBeenCalledWith('Cache-Control', 'no-store');
      expect(response.setHeader).toHaveBeenCalledWith(
        'Content-Type',
        'application/javascript; charset=utf-8',
      );
      const script = response.end.mock.calls[0][0];
      const config = result.configs.find((c) => c.matrix === `wps-${kind}`);
      expect(script).toContain(JSON.stringify(config.url));
      expect(script).toContain(JSON.stringify(config.fixturePath));
      for (const other of result.configs.filter((c) => c !== config))
        expect(script).not.toContain(JSON.stringify(other.url));
      const { app } = fixtureApp(config);
      const harness = bootstrapHarness(script, app);
      expect(harness.window.OnAddinLoad()).toBe(true);
      expect(harness.window.OnOpenHelper()).toBe(true);
      expect(app.CreateTaskPane).toHaveBeenCalledExactlyOnceWith(config.url);
    }
    for (const url of [
      '/wps/writer/',
      '/wps/sheet/index.html',
      '/wps/writer/ribbon.xml',
      '/wps/sheet/ribbon.xml',
      '/wps/writer/manifest.xml',
      '/wps/main.js',
      '/wps-native-test/writer/main.js',
      '/wps/writer/main.js.bak',
    ]) {
      const response = middlewareResponse();
      const next = vi.fn();
      middleware({ method: 'GET', url }, response, next);
      expect(next).toHaveBeenCalledOnce();
      expect(response.end).not.toHaveBeenCalled();
    }
    const instructions = await readFile(join(result.directory, 'WPS-ENTRY.md'), 'utf8');
    expect(instructions).toContain('无需改 publish.xml');
    expect(instructions).toContain('小帮手 → 帮我一下');
  });

  it('注册入口 HEAD 不返回 token 脚本，非读取请求拒绝；缺失或重复矩阵不能回退到普通入口', async () => {
    const result = await prepareExtendedCheck(await setup());
    const middleware = createWpsExtendedEntryMiddleware(result.configs);
    const head = middlewareResponse();
    const next = vi.fn();
    middleware({ method: 'HEAD', url: '/wps/writer/main.js' }, head, next);
    expect(head.end).toHaveBeenCalledExactlyOnceWith(undefined);
    const post = middlewareResponse();
    middleware({ method: 'POST', url: '/wps/sheet/main.js' }, post, next);
    expect(post.statusCode).toBe(405);
    expect(post.end).toHaveBeenCalledExactlyOnceWith();
    expect(next).not.toHaveBeenCalled();
    expect(() => createWpsExtendedEntryMiddleware(result.configs.slice(0, 3))).toThrow('唯一');
    expect(() => createWpsExtendedEntryMiddleware([...result.configs, result.configs[3]])).toThrow(
      '唯一',
    );
  });

  it.each(['text', 'sheet'])(
    '注册 %s 入口只轮询匹配样本并保留原回调，不自动执行 feature',
    async (kind) => {
      const result = await prepareExtendedCheck(await setup());
      const config = result.configs.find((c) => c.matrix === `wps-${kind}`);
      const { app, pane, source } = fixtureApp(config);
      const before = JSON.stringify(source);
      const harness = bootstrapHarness(
        wpsBootstrap(config, config.url, { registeredEntry: true }),
        app,
      );
      expect(harness.window.OnAddinLoad()).toBe(true);
      expect(harness.window.OnOpenHelper).toBe(harness.window.OnOpenExtendedCheck);
      expect(app.CreateTaskPane).not.toHaveBeenCalled();
      expect(harness.delayed[0].delay).toBe(1200);
      harness.delayed[0].callback();
      expect(harness.intervals[0].delay).toBe(700);
      harness.intervals[0].callback();
      expect(app.CreateTaskPane).toHaveBeenCalledExactlyOnceWith(config.url);
      expect(pane.Visible).toBe(true);
      expect(harness.clearInterval).toHaveBeenCalledExactlyOnceWith(1);
      expect(harness.window.OnOpenHelper()).toBe(true);
      expect(app.CreateTaskPane).toHaveBeenCalledOnce();
      expect(JSON.stringify(source)).toBe(before);
      source.Name = '普通用户文件.docx';
      pane.Visible = false;
      expect(harness.window.OnOpenHelper()).toBe(false);
      expect(app.CreateTaskPane).toHaveBeenCalledOnce();
      expect(pane.Visible).toBe(false);
    },
  );

  it.each(['text', 'sheet'])(
    '注册 %s 入口拒绝普通文件、同名其它路径、其它轮次/矩阵和不完整路径',
    async (kind) => {
      const result = await prepareExtendedCheck(await setup());
      const config = result.configs.find((c) => c.matrix === `wps-${kind}`);
      for (const change of [
        { Name: '普通文件.docx' },
        { FullName: join(dirname(config.fixturePath), 'other', config.fileName) },
        { FullName: config.fixturePath.replace(config.run, '0'.repeat(32)) },
        { FullName: config.fixturePath.replace(`wps-${kind}`, `office-${kind}`) },
        { FullName: config.fileName },
        { FullName: '' },
        { FullName: `${config.fixturePath}.bak` },
      ]) {
        const { app, source } = fixtureApp(config);
        Object.assign(source, change);
        const harness = bootstrapHarness(
          wpsBootstrap(config, config.url, { registeredEntry: true }),
          app,
        );
        expect(harness.window.OnOpenHelper()).toBe(false);
        harness.delayed[0].callback();
        for (let i = 0; i < 90; i++) harness.intervals[0].callback();
        expect(app.CreateTaskPane).not.toHaveBeenCalled();
        expect(harness.clearInterval).toHaveBeenCalledExactlyOnceWith(1);
        expect(harness.status.textContent).toContain('拒绝');
      }
    },
  );

  it('文件路径匹配后仍拒绝正文/RUN/MATRIX/矩阵篡改和错误工作表', async () => {
    const result = await prepareExtendedCheck(await setup());
    for (const config of result.configs.filter((c) => c.host === 'wps')) {
      for (const mismatch of ['content', 'run', 'matrix']) {
        const { app, source } = fixtureApp(config);
        if (config.kind === 'text') {
          source.Content.Text =
            mismatch === 'content'
              ? `${source.Content.Text}extra`
              : source.Content.Text.replace(
                  mismatch === 'run' ? config.run : config.matrix,
                  'wrong',
                );
        } else {
          const range = app.ActiveSheet.Range;
          app.ActiveSheet.Range = (address) => {
            const value = range(address);
            if (mismatch === 'content' && address === 'B3:F10') value.Value2[1][0] = 'modified';
            if (mismatch !== 'content' && address === 'A1:C1')
              value.Value2[0][mismatch === 'run' ? 1 : 2] = 'wrong';
            return value;
          };
        }
        const harness = bootstrapHarness(
          wpsBootstrap(config, config.url, { registeredEntry: true }),
          app,
        );
        expect(harness.window.OnOpenHelper()).toBe(false);
        harness.delayed[0].callback();
        harness.intervals[0].callback();
        expect(app.CreateTaskPane).not.toHaveBeenCalled();
      }
    }
    const config = result.configs.find((c) => c.matrix === 'wps-sheet');
    const { app } = fixtureApp(config);
    app.ActiveSheet.Name = 'ScopeSentinel';
    const harness = bootstrapHarness(
      wpsBootstrap(config, config.url, { registeredEntry: true }),
      app,
    );
    expect(harness.window.OnOpenHelper()).toBe(false);
    expect(app.CreateTaskPane).not.toHaveBeenCalled();
  });

  it('短轮询等待真实桥接/专用文件出现，WPS 本体桥接及数字 pane ID 可用，路径按 Windows 大小写比较', async () => {
    const result = await prepareExtendedCheck(await setup());
    const config = result.configs.find((c) => c.matrix === 'wps-text');
    const { app, source, pane } = fixtureApp(config);
    const harness = bootstrapHarness(wpsBootstrap(config, config.url, { registeredEntry: true }));
    harness.delayed[0].callback();
    harness.intervals[0].callback();
    expect(app.CreateTaskPane).not.toHaveBeenCalled();
    source.FullName = source.FullName.replaceAll('\\', '/').toUpperCase();
    app.CreateTaskPane.mockReturnValue(7);
    app.GetTaskPane = vi.fn(() => pane);
    harness.window.wps = app;
    harness.intervals[0].callback();
    expect(app.CreateTaskPane).toHaveBeenCalledExactlyOnceWith(config.url);
    expect(app.GetTaskPane).toHaveBeenCalledExactlyOnceWith(7);
    expect(pane.Visible).toBe(true);
  });

  it('缺少宿主接口明确 unsupported；创建异常或无 GetTaskPane 不重复创建、不伪报 passed', async () => {
    const result = await prepareExtendedCheck(await setup());
    const config = result.configs.find((c) => c.matrix === 'wps-text');
    const missing = fixtureApp(config);
    delete missing.app.CreateTaskPane;
    const unavailable = bootstrapHarness(
      wpsBootstrap(config, config.url, { registeredEntry: true }),
      missing.app,
    );
    expect(unavailable.window.OnOpenHelper()).toBe(false);
    expect(unavailable.status.textContent).toContain('unsupported');
    for (const fail of [
      () => 7,
      () => {
        throw new Error('no native pane');
      },
    ]) {
      const { app } = fixtureApp(config);
      app.CreateTaskPane.mockImplementation(fail);
      const harness = bootstrapHarness(
        wpsBootstrap(config, config.url, { registeredEntry: true }),
        app,
      );
      harness.delayed[0].callback();
      harness.intervals[0].callback();
      expect(harness.window.OnOpenHelper()).toBe(false);
      expect(app.CreateTaskPane).toHaveBeenCalledOnce();
      expect(harness.clearInterval).toHaveBeenCalledExactlyOnceWith(1);
      expect(harness.status.textContent).toContain('unsupported');
      expect(harness.status.textContent).not.toContain('passed');
    }
  });
});
