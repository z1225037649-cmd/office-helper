import type { DocumentKind, HostAdapter } from '../core/contracts';
import { HelperError } from '../core/errors';
import { MockAdapter } from './mock/adapter';

/** 显式宿主连接失败时不能悄悄切到演示模式。 */
export async function createAdapter(
  search: string,
  demoKind: DocumentKind = 'text',
): Promise<HostAdapter> {
  const params = new URLSearchParams(search);
  const host = params.get('host') ?? 'mock';
  if (host === 'mock') return new MockAdapter(demoKind);
  if (host === 'office') {
    await loadOfficeScript();
    const { connectOffice } = await import('./office/adapter');
    return connectOffice();
  }
  if (host === 'wps') {
    const kind = params.get('kind');
    if (kind !== 'text' && kind !== 'sheet')
      throw new HelperError('HOST_CONNECTION', 'WPS 入口未指定文字或表格类型。');
    const { WpsAdapter } = await import('./wps/adapter');
    const adapter = new WpsAdapter(kind);
    await adapter.readSelection();
    return adapter;
  }
  throw new HelperError('HOST_CONNECTION', '无法识别这个插件入口。');
}

let officeScript: Promise<void> | undefined;
function loadOfficeScript(): Promise<void> {
  if (typeof Office !== 'undefined') return Promise.resolve();
  officeScript ??= new Promise<void>((resolve, reject) => {
    const script = document.createElement('script');
    const fail = () => {
      script.remove();
      reject(new HelperError('HOST_CONNECTION', 'Office 连接组件未能加载，请检查网络后重试。'));
    };
    const timer = setTimeout(fail, 12000);
    script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    script.onload = () => {
      clearTimeout(timer);
      resolve();
    };
    script.onerror = () => {
      clearTimeout(timer);
      fail();
    };
    document.head.appendChild(script);
  }).catch((error) => {
    officeScript = undefined;
    throw error;
  });
  return officeScript;
}
