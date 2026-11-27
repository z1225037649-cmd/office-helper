import type { CellValue, DocumentKind } from '../../core/contracts';
import { HelperError } from '../../core/errors';

export interface WpsSelection {
  Text?: string;
  Start?: number;
  End?: number;
  Font?: { Size?: number; Bold?: number | boolean };
  Rows?: { Count: number };
  Columns?: { Count: number };
  Address?: string | (() => string);
  Value2?: CellValue | CellValue[][];
  Areas?: { Count: number };
  Cells?: { Item(row: number, column: number): WpsSheetCell };
  Range?: WpsRange;
}

/** Minimal native cell handle; it stays inside the WPS adapter. */
export interface WpsSheetCell {
  Value2?: CellValue;
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
  Font: { Size: number; Bold?: number | boolean };
  Document: WpsDocument;
  Duplicate: WpsRange;
  StoryType?: number;
  Tables?: { Count: number };
  Paragraphs?: { Count: number; Item(index: number): { Range: WpsRange } };
  ParagraphFormat?: {
    Alignment?: number;
    LineSpacing?: number;
    LineSpacingRule?: number;
    FirstLineIndent?: number;
    CharacterUnitFirstLineIndent?: number;
    LeftIndent?: number;
    RightIndent?: number;
  };
  ListFormat?: { ListType?: number };
}

export interface WpsApplication {
  Enum?: {
    wdAlignParagraphLeft?: number;
    wdAlignParagraphCenter?: number;
    wdAlignParagraphRight?: number;
    wdAlignParagraphJustify?: number;
    wdMainTextStory?: number;
    wdListNoNumbering?: number;
    WdParagraphAlignment?: Record<string, number>;
    WdStoryType?: Record<string, number>;
    WdListType?: Record<string, number>;
    wdLineSpaceSingle?: number;
    wdLineSpace1pt5?: number;
    wdLineSpaceDouble?: number;
    wdLineSpaceMultiple?: number;
    WdLineSpacing?: Record<string, number>;
  };
  ActiveDocument?: WpsDocument;
  ActiveWorkbook?: { Name: string; FullName?: string };
  ActiveSheet?: { Name: string };
  Selection?: WpsSelection;
  WorksheetFunction?: { IsError(cell: WpsSheetCell): boolean | number };
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
