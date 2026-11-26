import type { CellValue, DocumentKind } from '../../core/contracts';
import { HelperError } from '../../core/errors';

export interface WpsSelection {
  Text?: string;
  Start?: number;
  End?: number;
  Font?: { Size?: number };
  Rows?: { Count: number };
  Columns?: { Count: number };
  Address?: string | (() => string);
  Value2?: CellValue | CellValue[][];
  Range?: WpsRange;
}

export interface WpsDocument {
  Name: string;
  FullName?: string;
  ReadOnly?: boolean;
  ProtectionType?: number;
}

export interface WpsRange {
  Text: string;
  Start: number;
  End: number;
  Font: { Size: number };
  Document: WpsDocument;
  Duplicate: WpsRange;
}

export interface WpsApplication {
  ActiveDocument?: WpsDocument;
  ActiveWorkbook?: { Name: string; FullName?: string };
  ActiveSheet?: { Name: string };
  Selection?: WpsSelection;
  ActiveWindow?: { Hwnd?: number };
}

declare global {
  interface Window {
    Application?: WpsApplication;
    wps?: {
      Application?: WpsApplication;
      WpsApplication?: () => WpsApplication;
      EtApplication?: () => WpsApplication;
    };
  }
}

export function getWpsApplication(kind: DocumentKind): WpsApplication {
  const api = window.wps;
  const app =
    window.Application ??
    api?.Application ??
    (kind === 'text' ? api?.WpsApplication?.() : api?.EtApplication?.());
  if (!app)
    throw new HelperError('HOST_CONNECTION', '没有连接到 WPS，请从 WPS 内的“小帮手”按钮打开。');
  return app;
}
