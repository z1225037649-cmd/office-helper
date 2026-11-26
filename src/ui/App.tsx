import { useEffect, useRef, useState } from 'react';
import {
  ArrowRight,
  Check,
  ChevronRight,
  CircleHelp,
  FileText,
  History,
  Leaf,
  LoaderCircle,
  ScanText,
  ShieldCheck,
  Sigma,
  Sparkles,
  Table2,
  Type,
  Undo2,
  X,
} from 'lucide-react';
import type { AppRuntime } from '../app/runtime';
import type { DocumentKind, PreparedTask, SelectionSnapshot } from '../core/contracts';
import { userMessage } from '../core/errors';
import { DemoDocument } from './DemoDocument';

const featureIcons = { text: Type, sum: Sigma, inspect: ScanText };

export function App({
  runtime,
  onDemoSwitch,
}: {
  runtime: AppRuntime;
  onDemoSwitch: (kind: DocumentKind) => void;
}) {
  const { host, registry, runner, intent } = runtime;
  const [selection, setSelection] = useState<SelectionSnapshot>();
  const [preview, setPreview] = useState<PreparedTask>();
  const [result, setResult] = useState('');
  const [error, setError] = useState('');
  const [input, setInput] = useState('');
  const [busy, setBusy] = useState(false);
  const [large, setLarge] = useState(false);
  const [hint, setHint] = useState('');
  const previewRef = useRef<HTMLHeadingElement>(null);
  const isDemo = host.info.id === 'mock';
  const visibleFeatures = registry
    .list()
    .filter((feature) => feature.kinds.includes(host.info.kind));

  useEffect(() => {
    let active = true;
    host
      .readSelection()
      .then((value) => {
        if (active) setSelection(value);
      })
      .catch((reason) => {
        if (active) setError(userMessage(reason));
      });
    return () => {
      active = false;
    };
  }, [host]);

  useEffect(() => {
    if (preview) previewRef.current?.focus();
  }, [preview]);

  async function perform(action: () => Promise<void>) {
    setBusy(true);
    setError('');
    setHint('');
    try {
      await action();
    } catch (reason) {
      setError(userMessage(reason));
    } finally {
      setBusy(false);
    }
  }

  function prepare(id: string) {
    return perform(async () => {
      if (preview) runner.cancel(preview.id);
      setPreview(undefined);
      setResult('');
      const next = await runner.prepare(id);
      setPreview(next);
      setSelection(await host.readSelection());
    });
  }

  function confirm() {
    if (!preview) return;
    const id = preview.id;
    setPreview(undefined);
    void perform(async () => {
      const outcome = await runner.execute(id);
      setResult(outcome.message);
      setSelection(await host.readSelection());
    });
  }

  async function submitIntent() {
    await perform(async () => {
      if (preview) runner.cancel(preview.id);
      setPreview(undefined);
      setResult('');
      const available = visibleFeatures.filter(
        (feature) => !registry.unavailableReason(feature, host.info),
      );
      const response = await intent.resolve(input, available);
      if (response.type === 'match') {
        setPreview(await runner.prepare(response.featureId, response.parameters));
        setSelection(await host.readSelection());
      } else if (response.type === 'choices') {
        setHint('我们一步一步来，请先点下面一个功能。');
      } else if (response.type === 'clarify') {
        setHint(response.message);
      } else {
        setHint('这件事暂时还不会做。你可以先试试下面的功能。');
      }
    });
  }

  return (
    <div className={`app ${isDemo ? 'demo-mode' : 'native-mode'} ${large ? 'large-text' : ''}`}>
      {isDemo && (
        <header className="site-header">
          <a className="brand" href="/">
            <span className="brand-icon">
              <Leaf size={24} />
            </span>
            <span>
              小帮手<small>办公，轻松一点。</small>
            </span>
          </a>
          <div className="demo-badge">
            <span />
            交互演示 · 不修改真实文件
          </div>
        </header>
      )}
      <main className="app-layout">
        {isDemo && (
          <DemoDocument
            selection={selection}
            kind={host.info.kind}
            disabled={busy}
            onSwitch={onDemoSwitch}
          />
        )}
        <aside className="assistant-panel" aria-label="办公小帮手">
          <div className="assistant-top">
            <div>
              <span className="assistant-symbol">
                <Sparkles size={20} />
              </span>
              <strong>小帮手</strong>
              <span className="online-dot" />
            </div>
            <button
              className="icon-button text-size-button"
              aria-label={large ? '恢复界面字号' : '放大界面字号'}
              aria-pressed={large}
              onClick={() => setLarge(!large)}
            >
              <Type size={19} />
              <span>字号</span>
            </button>
          </div>
          <div className="assistant-body">
            <div className="welcome">
              <span className="eyebrow">一步一步，轻松完成</span>
              <h1>今天想做些什么？</h1>
              <p>选一个小任务，剩下的交给我。</p>
            </div>
            <div className="file-context">
              {host.info.kind === 'text' ? <FileText size={23} /> : <Table2 size={23} />}
              <div>
                <strong>{selection?.documentName ?? '正在读取文件…'}</strong>
                <span>{selection?.targetLabel ?? host.info.label}</span>
              </div>
              <span className="context-label">{isDemo ? '示例' : '已连接'}</span>
            </div>
            {!isDemo && host.info.notice && <p className="host-notice">{host.info.notice}</p>}
            <form
              className="request-form"
              onSubmit={(event) => {
                event.preventDefault();
                void submitIntent();
              }}
            >
              <label htmlFor="request">告诉我，你想怎么处理？</label>
              <div className="request-input">
                <input
                  id="request"
                  value={input}
                  disabled={busy}
                  onChange={(event) => setInput(event.target.value)}
                  placeholder={
                    host.info.kind === 'text' ? '例如：把字放大一点' : '例如：算一下合计'
                  }
                  maxLength={300}
                  autoComplete="off"
                />
                <button type="submit" disabled={busy || !input.trim()} aria-label="理解这句话">
                  <ArrowRight size={21} />
                </button>
              </div>
            </form>
            {hint && (
              <p className="info-message" role="status">
                <CircleHelp size={18} />
                {hint}
              </p>
            )}
            <div className="section-label">
              <span>也可以直接点这里</span>
              <span className="line" />
            </div>
            <div className="feature-list">
              {visibleFeatures.map((feature) => {
                const Icon = featureIcons[feature.icon];
                const reason = registry.unavailableReason(feature, host.info);
                return (
                  <button
                    key={feature.id}
                    className="feature-card"
                    disabled={busy || !!reason}
                    title={reason}
                    onClick={() => void prepare(feature.id)}
                  >
                    <span className={`feature-icon ${feature.icon}`}>
                      <Icon size={25} />
                    </span>
                    <span>
                      <strong>{feature.title}</strong>
                      <small>{reason ?? feature.description}</small>
                    </span>
                    <ChevronRight size={18} className="feature-arrow" />
                  </button>
                );
              })}
            </div>
            {busy && (
              <div className="working" role="status">
                <LoaderCircle className="spin" size={19} />
                正在处理，请稍等…
              </div>
            )}
            {preview && (
              <section className="preview-card" aria-label="操作预览">
                <span className="eyebrow">先看一眼，再继续</span>
                <h2 tabIndex={-1} ref={previewRef}>
                  {preview.title}
                </h2>
                <p className="preview-target">{preview.target}</p>
                <p>{preview.summary}</p>
                <div className="preview-buttons">
                  <button className="primary-button" disabled={busy} onClick={confirm}>
                    {preview.changesDocument ? '确认修改' : '查看结果'}
                    <ArrowRight size={17} />
                  </button>
                  <button
                    className="secondary-button"
                    disabled={busy}
                    onClick={() => {
                      runner.cancel(preview.id);
                      setPreview(undefined);
                    }}
                  >
                    先不做
                  </button>
                </div>
              </section>
            )}
            {result && (
              <div className="result-card" role="status">
                <Check size={22} />
                <div>
                  <strong>完成了</strong>
                  <p>{result}</p>
                </div>
                <button className="icon-button" aria-label="收起结果" onClick={() => setResult('')}>
                  <X size={16} />
                </button>
              </div>
            )}
            {error && (
              <div className="error-card" role="alert">
                <CircleHelp size={20} />
                <p>{error}</p>
              </div>
            )}
            <button
              className="undo-button"
              disabled={busy || !runner.canUndo}
              onClick={() =>
                void perform(async () => {
                  await runner.undo();
                  setPreview(undefined);
                  setResult('已经恢复到上一次修改之前。');
                  setSelection(await host.readSelection());
                })
              }
            >
              <Undo2 size={18} />
              恢复上一次修改
            </button>
            <details className="activity">
              <summary>
                <History size={16} />
                刚才做过的事<span>{runner.history.length}</span>
              </summary>
              {runner.history.length ? (
                <ol>
                  {runner.history.map((entry) => (
                    <li key={entry.id}>
                      <span>{entry.title}</span>
                      <time>
                        {new Date(entry.time).toLocaleTimeString('zh-CN', {
                          hour: '2-digit',
                          minute: '2-digit',
                        })}
                      </time>
                    </li>
                  ))}
                </ol>
              ) : (
                <p>完成任务后，会在这里留下记录。</p>
              )}
            </details>
          </div>
          <footer className="assistant-footer">
            <ShieldCheck size={16} />
            <span>{isDemo ? '放心试试，每次修改都可以恢复' : '操作前先确认选中的范围'}</span>
          </footer>
        </aside>
      </main>
      {isDemo && (
        <footer className="site-footer">
          <span>小帮手 · 让日常办公更简单</span>
          <span>文字与表格，共用一个熟悉的助手</span>
        </footer>
      )}
    </div>
  );
}
