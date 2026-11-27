import { createAdapter } from '../adapters';
import { createExtendedObserver } from './extended-check-host';
import { ExtendedCheckSession } from './extended-check-session';
import {
  assertExtendedConfig,
  featuresFor,
  type ExtendedConfig,
  type ExtendedFeature,
} from './extended-check-spec';

/** Only the preparation script's development middleware serves this entry. */
async function start() {
  if (!import.meta.env.DEV) throw new Error('扩展诊断只允许本地开发服务器。');
  const params = new URLSearchParams(location.search);
  const token = params.get('token') ?? '';
  if (
    !/^[a-f0-9]{48}$/u.test(token) ||
    !/^\/__extended\/[a-f0-9]{32}\/(office|wps)-(text|sheet)\/$/u.test(location.pathname)
  )
    throw new Error('缺少隔离入口或本地 token。');
  const root = document.getElementById('extended-root');
  if (!root) throw new Error('缺少诊断页面。');
  const request = async (operation: string): Promise<unknown> => {
    const response = await fetch(`${location.pathname}${operation}`, {
      method: operation === 'config' ? 'GET' : 'POST',
      headers: { 'X-Helper-Test': token },
      cache: 'no-store',
    });
    if (!response.ok) throw new Error(`隔离本地端点拒绝：${await response.text()}`);
    return response.json();
  };
  const config = (await request('config')) as ExtendedConfig;
  assertExtendedConfig(config);
  if (
    params.get('host') !== config.host ||
    params.get('kind') !== config.kind ||
    location.pathname !== `/__extended/${config.run}/${config.matrix}/` ||
    location.origin !==
      (config.host === 'office' ? 'https://localhost:3443' : 'http://127.0.0.1:5173')
  )
    throw new Error('入口 URL 与矩阵配置不匹配。');
  const status = document.createElement('pre');
  status.style.whiteSpace = 'pre-wrap';
  const info = document.createElement('p');
  info.textContent = `${config.matrix}；样本：${config.fixturePath}；唯一结果目录：${config.resultDirectory}。连接前请按 CHECKLIST.md 选中全部正文或 ExtendedData!B3:F10。`;
  root.append(info);
  let session: ExtendedCheckSession | undefined;
  const report = async (stage: string, details: Record<string, unknown>) => {
    const response = await fetch('/__native-report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-Helper-Test': token },
      body: JSON.stringify({
        suite: 'extended',
        run: config.run,
        runId: config.run,
        matrix: config.matrix,
        host: config.host,
        stage,
        ...details,
      }),
    });
    if (!response.ok)
      throw new Error(`本地报告拒绝（${response.status}）；未获占用锁或服务未启用。`);
    status.textContent = `${stage}\n${JSON.stringify(details, null, 2)}\n\n${JSON.stringify(session?.state, null, 2) ?? ''}`;
  };
  const controls: { button: HTMLButtonElement; enabled: () => boolean }[] = [];
  let running = false;
  const refresh = () => {
    for (const control of controls)
      control.button.disabled = running || !!session?.state.stopped || !control.enabled();
    feature.disabled = running || !!session?.state.pending || !!session?.state.undo;
  };
  const button = (
    label: string,
    operation: () => Promise<unknown>,
    needsSession = true,
    enabled = () => !!session?.state.claimed,
  ) => {
    const element = document.createElement('button');
    element.textContent = label;
    element.style.cssText = 'display:block;margin:8px 0;padding:8px;max-width:100%;';
    element.onclick = () => {
      if (running) return;
      running = true;
      refresh();
      void (async () => {
        try {
          if (needsSession) {
            if (!session) throw new Error('请先连接并核验专用样本。');
            await session.action(operation);
          } else await operation();
          if (session)
            status.textContent += `\n\n当前状态\n${JSON.stringify(session.state, null, 2)}`;
          renderResults();
        } catch (error) {
          status.textContent += `\n停止：${error instanceof Error ? error.message : String(error)}`;
        } finally {
          running = false;
          refresh();
        }
      })();
    };
    controls.push({ button: element, enabled });
    root.append(element);
  };
  button(
    '1. 连接并核验专用样本（只读；领取本轮单元）',
    async () => {
      if (session) throw new Error('本窗格已连接，请继续本轮。');
      const host = await createAdapter(location.search);
      try {
        const observer = createExtendedObserver(host, config);
        const candidate = new ExtendedCheckSession(host, config, observer, report, request);
        await candidate.initialize();
        session = candidate;
      } catch (error) {
        host.dispose();
        await report('extended-case', {
          check: 'connect',
          status: 'unsupported',
          reason: error instanceof Error ? error.message : String(error),
        });
        throw error;
      }
    },
    false,
    () => !session,
  );
  const feature = document.createElement('select');
  for (const id of featuresFor(config.kind)) {
    const option = document.createElement('option');
    option.value = id;
    option.textContent = id;
    feature.append(option);
  }
  root.append(feature);
  const pending = () => !!session?.state.pending;
  const undo = () => !!session?.state.undo;
  button(
    '2. 只预览所选功能',
    () => session!.prepare(feature.value as ExtendedFeature),
    true,
    () => !!session?.state.claimed && !pending() && !undo(),
  );
  button('3a. 取消并验证无写入', () => session!.cancel(), true, pending);
  button(
    '3b. 新预览替换旧预览，验证 EXPIRED_TASK（不写入）',
    () => session!.supersede(),
    true,
    pending,
  );
  button(
    '3c. 已用原生 UI 移动选区：验证 STALE_SELECTION（不写入）',
    () => session!.stale(),
    true,
    pending,
  );
  button(
    '4. 确认执行当前预览（真实 runner；导出每轮仅一次）',
    () => session!.confirm(),
    true,
    pending,
  );
  button('5. 恢复本轮格式/打印设置并逐项核对', () => session!.undo(), true, undo);
  button('5b. 已用原生 UI 改动本项属性：验证恢复冲突', () => session!.conflict(), true, undo);
  const results = document.createElement('div');
  root.append(results);
  const renderResults = () => {
    results.replaceChildren();
    for (const finding of session?.state.findings ?? []) {
      const item = document.createElement('button');
      item.textContent = `定位 ${finding.id}：${finding.label}`;
      item.onclick = () => runResult(() => session!.locate(finding.id));
      results.append(item);
    }
    for (const action of session?.state.result?.actions ?? []) {
      const item = document.createElement('button');
      item.textContent = action.label;
      item.onclick = () => runResult(() => session!.openResult(action.id));
      results.append(item);
    }
  };
  const runResult = (operation: () => Promise<unknown>) => {
    if (running || session?.state.stopped) return;
    running = true;
    refresh();
    void session!
      .action(operation)
      .catch((error) => {
        status.textContent += `\n${String(error)}`;
      })
      .finally(() => {
        running = false;
        refresh();
      });
  };
  button('只读核对专用目录中的副本（可用于原生 UI 手工保存后）', () =>
    session!.verifyOutput(false),
  );
  button('只读核对专用目录中的 PDF（仍须人工核对全范围）', () => session!.verifyOutput(true));
  const manual = document.createElement('select');
  for (const id of [
    'native-dialog-cancel',
    'native-collision-cancel',
    'native-wrong-folder-cancel',
    'print-preview',
    'output-scope-location',
    'notice-new-window',
    'switch-dedicated-file-stale',
    'records-data-stale',
  ]) {
    const option = document.createElement('option');
    option.value = id;
    option.textContent = id;
    manual.append(option);
  }
  const evidence = document.createElement('textarea');
  evidence.placeholder = '实际软件版本、正常 UI 操作结果、证据位置；人工证据不提升自动项状态。';
  evidence.style.cssText = 'display:block;width:95%;min-height:90px';
  root.append(manual, evidence);
  button('记录所选人工项（仅 manual）', () => session!.manual(manual.value, evidence.value));
  button('输出本轮证据汇总（不自动宣称总验收通过）', () => session!.summary());
  root.append(status);
  status.textContent = '尚未连接或操作任何宿主。请先阅读本轮 CHECKLIST.md。';
  refresh();
}

void start().catch((error) => {
  const status = document.createElement('pre');
  status.textContent = `入口停止：${error instanceof Error ? error.message : String(error)}`;
  document.body.append(status);
});
