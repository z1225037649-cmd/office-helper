import { Component, useEffect, useState, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import { Leaf, RefreshCw } from 'lucide-react';
import { createAdapter } from './adapters';
import { createRuntime, type AppRuntime } from './app/runtime';
import type { DocumentKind, HostAdapter } from './core/contracts';
import { userMessage } from './core/errors';
import { App } from './ui/App';
import './ui/styles.css';

function Root() {
  const [kind, setKind] = useState<DocumentKind>('text');
  const [runtime, setRuntime] = useState<AppRuntime>();
  const [error, setError] = useState('');
  const [attempt, setAttempt] = useState(0);
  useEffect(() => {
    let disposed = false;
    let adapter: HostAdapter | undefined;
    setRuntime(undefined);
    setError('');
    void createAdapter(location.search, kind)
      .then((host) => {
        if (disposed) {
          host.dispose();
          return;
        }
        adapter = host;
        setRuntime(createRuntime(host));
      })
      .catch((reason) => {
        if (!disposed) setError(userMessage(reason));
      });
    return () => {
      disposed = true;
      adapter?.dispose();
    };
  }, [kind, attempt]);

  if (error)
    return (
      <div className="connection-screen" role="alert">
        <Leaf size={38} />
        <h1>暂时没有连接上</h1>
        <p>{error}</p>
        <button className="primary-button" onClick={() => setAttempt(attempt + 1)}>
          <RefreshCw size={18} />
          重新连接
        </button>
      </div>
    );
  if (!runtime)
    return (
      <div className="connection-screen" role="status">
        <Leaf size={38} />
        <p>正在打开小帮手…</p>
      </div>
    );
  return (
    <App
      key={`${runtime.host.info.id}-${kind}-${attempt}`}
      runtime={runtime}
      onDemoSwitch={setKind}
    />
  );
}

class ErrorBoundary extends Component<{ children: ReactNode }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() {
    return { failed: true };
  }
  render() {
    return this.state.failed ? (
      <div className="connection-screen" role="alert">
        <h1>页面暂时没有打开</h1>
        <p>请重新打开助手，再试一次。</p>
        <button className="primary-button" onClick={() => location.reload()}>
          重新打开
        </button>
      </div>
    ) : (
      this.props.children
    );
  }
}

createRoot(document.getElementById('root')!).render(
  <ErrorBoundary>
    <Root />
  </ErrorBoundary>,
);
