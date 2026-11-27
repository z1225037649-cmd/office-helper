import { useEffect, useRef, useState } from 'react';
import {
  ArrowRight,
  Check,
  CircleHelp,
  FileText,
  History,
  Leaf,
  LoaderCircle,
  ShieldCheck,
  Sparkles,
  Table2,
  Type,
  Undo2,
  X,
} from 'lucide-react';
import type { AppRuntime } from '../app/runtime';
import type {
  DocumentKind,
  FeatureParameters,
  PreparedTask,
  SelectionSnapshot,
  TaskResult,
} from '../core/contracts';
import { userMessage } from '../core/errors';
import { selectionKey } from '../core/selection';
import { DemoDocument } from './DemoDocument';
import { FeatureBrowser } from './FeatureBrowser';
import { FeatureForms, type FeatureFormContext } from './FeatureForms';
import { MockAdapter } from '../adapters/mock/adapter';
import { SheetBatchTool } from './SheetBatchTool';

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
  const [result, setResult] = useState<TaskResult>();
  const [formContext, setFormContext] = useState<FeatureFormContext>();
  const [formReview, setFormReview] = useState<{
    context: FeatureFormContext;
    parameters: FeatureParameters;
  }>();
  const [error, setError] = useState('');
  const [input, setInput] = useState('');
  const [busy, setBusy] = useState(false);
  const [large, setLarge] = useState(false);
  const [batchToolOpen, setBatchToolOpen] = useState(false);
  const [hint, setHint] = useState('');
  const previewRef = useRef<HTMLHeadingElement>(null);
  const previewCardRef = useRef<HTMLElement>(null);
  const hintRef = useRef<HTMLParagraphElement>(null);
  const resultRef = useRef<HTMLDivElement>(null);
  const errorRef = useRef<HTMLDivElement>(null);
  const formRef = useRef<HTMLDivElement>(null);
  const isDemo = host.info.id === 'mock';
  const visibleFeatures = registry
    .list()
    .filter((feature) => feature.kinds.includes(host.info.kind));

  useEffect(() => {
    let active = true;
    host
      .readSelection(
        host.info.id === 'mock'
          ? { includeSheetLayout: true, includePrintSettings: true }
          : undefined,
      )
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
    if (busy) return;
    const feedback = error
      ? errorRef.current
      : preview
        ? previewCardRef.current
        : formContext
          ? formRef.current
          : hint
            ? hintRef.current
            : result
              ? resultRef.current
              : null;
    if (preview && !error) previewRef.current?.focus({ preventScroll: true });
    if (formContext && !preview && !error)
      formRef.current?.querySelector<HTMLInputElement>('input')?.focus({ preventScroll: true });
    feedback?.scrollIntoView({ block: 'nearest' });
  }, [busy, error, formContext, hint, preview, result]);

  async function refreshSelection() {
    setSelection(
      await host.readSelection(
        isDemo ? { includeSheetLayout: true, includePrintSettings: true } : undefined,
      ),
    );
  }

  async function openForm(id: string) {
    const feature = registry.get(id);
    const current = await host.readSelection(feature.readOptions);
    setFormContext({ feature, selection: current });
    await refreshSelection();
  }

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
      setBatchToolOpen(false);
      if (preview) runner.cancel(preview.id);
      setPreview(undefined);
      setResult(undefined);
      setFormContext(undefined);
      setFormReview(undefined);
      const feature = registry.get(id);
      if (feature.form) {
        await openForm(id);
        return;
      }
      const prompt = feature.inputPrompt;
      if (prompt) {
        setInput(prompt.prefix);
        setHint(prompt.message);
        document.getElementById('request')?.focus();
        return;
      }
      const next = await runner.prepare(id);
      setPreview(next);
      await refreshSelection();
    });
  }

  function confirm() {
    if (!preview) return;
    const id = preview.id;
    setPreview(undefined);
    void perform(async () => {
      const outcome = await runner.execute(id);
      setFormReview(undefined);
      setResult(outcome);
      await refreshSelection();
    });
  }

  function reopenFailedForm() {
    if (!formReview) return;
    const draft = formReview;
    void perform(async () => {
      const current = await host.readSelection(draft.context.feature.readOptions);
      const changed = selectionKey(current) !== selectionKey(draft.context.selection);
      const parameters = structuredClone(draft.parameters);
      if (changed && parameters.duplicates) {
        parameters.duplicates.keyColumns = [];
        parameters.duplicates.requiredColumns = [];
      }
      const context = { feature: draft.context.feature, selection: current };
      setFormReview({ context, parameters });
      setFormContext(context);
      await refreshSelection();
      if (changed)
        setHint(
          parameters.duplicates
            ? '当前文件或范围已改变，请重新选择判重列和必填列，再核对预览。'
            : '已保留填写内容，请核对当前文件和范围后重新预览。',
        );
    });
  }

  async function submitIntent() {
    await perform(async () => {
      if (preview) runner.cancel(preview.id);
      setPreview(undefined);
      setResult(undefined);
      setFormContext(undefined);
      setFormReview(undefined);
      const available = visibleFeatures.filter(
        (feature) => !registry.unavailableReason(feature, host.info),
      );
      const response = await intent.resolve(input, available);
      if (response.type === 'match') {
        if (registry.get(response.featureId).form) {
          await openForm(response.featureId);
          return;
        }
        setPreview(await runner.prepare(response.featureId, response.parameters));
        await refreshSelection();
      } else if (response.type === 'choices') {
        setHint('我们一步一步来，请一次只说一项要求，并保留具体数值，例如“行距设为1.5倍”。');
      } else if (response.type === 'clarify') {
        setHint(response.message);
      } else {
        setHint('这件事暂时还不会做。你可以先试试下面的功能。');
      }
    });
  }

  function submitForm(parameters: FeatureParameters) {
    if (!formContext) return;
    const context = formContext;
    void perform(async () => {
      const next = await runner.prepare(context.feature.id, parameters, context.selection);
      setFormReview({ context, parameters: structuredClone(parameters) });
      setFormContext(undefined);
      setPreview(next);
      await refreshSelection();
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
            generatedNotice={host instanceof MockAdapter ? host.createdDocuments.at(-1) : undefined}
            locatedCell={host instanceof MockAdapter ? host.locatedCell : undefined}
          />
        )}
        <aside className="assistant-panel" aria-label="办公小帮手">
          <div className="assistant-top">
            <div>
              <span className="assistant-symbol">
                <Sparkles size={18} />
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
              <Type size={16} />
              <span>{large ? '标准字号' : '大字号'}</span>
            </button>
          </div>
          <div className="assistant-body">
            <div className="welcome">
              <h1>{host.info.kind === 'text' ? '文字工具' : '表格工具'}</h1>
              <p>选中内容，选择功能或输入一句话。</p>
            </div>
            <div className="file-context">
              {host.info.kind === 'text' ? <FileText size={18} /> : <Table2 size={18} />}
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
              <label htmlFor="request">想怎么处理？</label>
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
                  <ArrowRight size={18} />
                </button>
              </div>
            </form>
            {hint && (
              <p className="info-message" role="status" ref={hintRef}>
                <CircleHelp size={18} />
                {hint}
              </p>
            )}
            {busy && (
              <div className="working" role="status">
                <LoaderCircle className="spin" size={19} />
                正在处理，请稍等…
              </div>
            )}
            {formContext && (
              <div ref={formRef}>
                <FeatureForms
                  key={`${formContext.feature.id}-${formContext.selection.revision}`}
                  context={formContext}
                  initialParameters={formReview?.parameters}
                  busy={busy}
                  onSubmit={submitForm}
                  onCancel={() => {
                    setFormContext(undefined);
                    setFormReview(undefined);
                  }}
                />
              </div>
            )}
            {preview && (
              <section className="preview-card" aria-label="操作预览" ref={previewCardRef}>
                <span className="eyebrow">先看一眼，再继续</span>
                <h2 tabIndex={-1} ref={previewRef}>
                  {preview.title}
                </h2>
                <p className="preview-target">{preview.target}</p>
                <p>{preview.summary}</p>
                <div className="preview-buttons">
                  <button className="primary-button" disabled={busy} onClick={confirm}>
                    {preview.confirmLabel ?? (preview.changesDocument ? '确认修改' : '查看结果')}
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
                  {formReview && (
                    <button
                      className="secondary-button"
                      disabled={busy}
                      onClick={() => {
                        runner.cancel(preview.id);
                        setPreview(undefined);
                        setFormContext(formReview.context);
                      }}
                    >
                      返回修改
                    </button>
                  )}
                </div>
              </section>
            )}
            {result && (
              <div className="result-card" role="status" ref={resultRef}>
                <Check size={22} />
                <div>
                  <strong>
                    {result.outcome === 'guidance'
                      ? '请在原软件继续操作'
                      : result.outcome === 'cancelled'
                        ? '已取消'
                        : result.outcome === 'download'
                          ? '已交给浏览器下载'
                          : '完成了'}
                  </strong>
                  <p>{result.message}</p>
                  {result.location && <p>结果：{result.location}</p>}
                  {!!result.actions?.length && (
                    <div className="preview-buttons">
                      {result.actions.map((action) => (
                        <button
                          className="secondary-button"
                          type="button"
                          key={action.id}
                          disabled={busy}
                          onClick={() => void perform(() => runner.openResult(action.id))}
                        >
                          {action.label}
                        </button>
                      ))}
                    </div>
                  )}
                  {!!result.findings?.length && (
                    <ol className="finding-list">
                      {result.findings.map((finding) => (
                        <li key={finding.id}>
                          <span>
                            <strong>{finding.label}</strong>
                            {finding.sourceExcerpt && <q>{finding.sourceExcerpt}</q>}
                            <span>{finding.description}</span>
                          </span>
                          {host.locate && host.info.capabilities.includes('selection.locate') && (
                            <button
                              type="button"
                              className="secondary-button"
                              disabled={busy}
                              onClick={() =>
                                void perform(async () => {
                                  await runner.locate(finding.id);
                                  setHint(`已定位：${finding.label}`);
                                  await refreshSelection();
                                })
                              }
                            >
                              定位
                            </button>
                          )}
                        </li>
                      ))}
                    </ol>
                  )}
                </div>
                <button
                  className="icon-button"
                  aria-label="收起结果"
                  onClick={() => setResult(undefined)}
                >
                  <X size={16} />
                </button>
              </div>
            )}
            {error && (
              <div className="error-card" role="alert" ref={errorRef}>
                <CircleHelp size={20} />
                <div>
                  <p>{error}</p>
                  {formReview && !formContext && !preview && (
                    <button
                      type="button"
                      className="secondary-button"
                      disabled={busy}
                      onClick={reopenFailedForm}
                    >
                      返回修改已填内容
                    </button>
                  )}
                </div>
              </div>
            )}
            <FeatureBrowser
              features={visibleFeatures}
              busy={busy}
              unavailableReason={(feature) => registry.unavailableReason(feature, host.info)}
              onPrepare={(id) => void prepare(id)}
              onOpenBatchTool={() => {
                if (preview) runner.cancel(preview.id);
                setPreview(undefined);
                setResult(undefined);
                setFormContext(undefined);
                setFormReview(undefined);
                setBatchToolOpen(true);
              }}
            />
            {batchToolOpen && (
              <SheetBatchTool busy={busy} onClose={() => setBatchToolOpen(false)} />
            )}
            <button
              className="undo-button"
              disabled={busy || !runner.canUndo}
              onClick={() =>
                void perform(async () => {
                  await runner.undo();
                  setPreview(undefined);
                  setFormContext(undefined);
                  setFormReview(undefined);
                  setResult({ message: '已经恢复到上一次修改之前。', canUndo: false });
                  await refreshSelection();
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
            <span>{isDemo ? '操作前先预览，格式修改可恢复' : '操作前先确认选中的范围'}</span>
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
