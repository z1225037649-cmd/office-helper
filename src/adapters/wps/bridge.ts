import type { CellValue, DocumentKind } from '../../core/contracts';
import { HelperError } from '../../core/errors';

export interface WpsSelection {
  Text?: string;
  Start?: number;
  End?: number;
  Font?: { Size?: number; Bold?: number | boolean };
  Rows?: { Count: number };
  Columns?: { Count: number; AutoFit?(): unknown };
  Address?: string | (() => string);
  /** Native bridges may return variants or incomplete matrices; validate at the boundary. */
  Value2?: unknown;
  Areas?: { Count: number };
  Cells?: { Item(row: number, column: number): WpsSheetCell };
  Range?: WpsRange;
  Row?: number;
  Column?: number;
  Worksheet?: WpsWorksheet;
}

/** Minimal native cell handle; it stays inside the WPS adapter. */
export interface WpsSheetCell {
  Value2?: CellValue;
  ColumnWidth?: number;
  WrapText?: boolean | number | null;
  Font?: { Bold?: number | boolean };
  MergeCells?: boolean | number | null;
  Select?(): unknown;
}

export interface WpsSheetPdfSource {
  ExportAsFixedFormat?(
    type: number,
    fileName: string,
    quality?: number,
    includeDocProperties?: boolean,
    ignorePrintAreas?: boolean,
    from?: number,
    to?: number,
    openAfterPublish?: boolean,
    fixedFormatExtClassPtr?: unknown,
  ): unknown;
}

export interface WpsWorkbook extends WpsSheetPdfSource {
  Name: string;
  FullName?: string;
  ReadOnly?: boolean;
  /** Sheets includes chart sheets; Worksheets alone cannot prove whole-workbook scope. */
  Sheets?: { Count: number; Item(index: number): WpsSheet };
  Worksheets?: { Item(name: string): WpsWorksheet };
  SaveCopyAs?(fileName: string): unknown;
}

export interface WpsWriterPdfSelection extends WpsSelection {
  ExportAsFixedFormat?(
    fileName: string,
    exportFormat: number,
    openAfterExport: boolean,
    optimizeFor?: number,
    exportCurrentPage?: boolean,
  ): unknown;
}

export interface WpsPageSetup {
  PrintArea?: string;
  Orientation?: number;
  Zoom?: number | false;
  FitToPagesWide?: number | false;
  FitToPagesTall?: number | false;
}

/** Common read-only scope metadata shared by chart sheets and worksheets. */
export interface WpsSheet {
  Name: string;
  Parent?: WpsWorkbook;
  Visible?: number;
}

export interface WpsWorksheet extends WpsSheet, WpsSheetPdfSource {
  ProtectContents?: boolean;
  Range?(address: string): WpsSelection;
  PageSetup?: WpsPageSetup;
  Activate?(): unknown;
}

export interface WpsFileDialog {
  InitialFileName?: string;
  Title?: string;
  ButtonName?: string;
  Show?(): number;
  SelectedItems?: { Count: number; Item(index: number): string };
}

export interface WpsFileSystem {
  Exists?(absolutePath: string): unknown;
  readAsBinaryString?(absolutePath: string): unknown;
  ReadFile?(absolutePath: string): unknown;
  copyFileSync?(source: string, target: string): unknown;
  readdirSync?(absolutePath: string): unknown;
}

export interface WpsDocument {
  Name: string;
  FullName?: string;
  ReadOnly?: boolean;
  ProtectionType?: number;
  Content?: {
    Text: string;
    Document?: WpsDocument;
    Start?: number;
    End?: number;
    StoryType?: number;
    Style?: string | number;
    Tables?: { Count: number };
    Paragraphs?: WpsRange['Paragraphs'];
  };
  Saved?: boolean;
  ExportAsFixedFormat?(fileName: string, exportFormat: number, openAfterExport: boolean): unknown;
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
  Style?: string | number;
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
    wdStyleNormal?: number;
    wdStyleHeading1?: number;
    wdStyleHeading2?: number;
    wdStyleHeading3?: number;
    WdLineSpacing?: Record<string, number>;
    xlLandscape?: number;
    xlPortrait?: number;
    XlPageOrientation?: Record<string, number>;
    msoFileDialogSaveAs?: number;
    MsoFileDialogType?: Record<string, number>;
    xlTypePDF?: number;
    XlFixedFormatType?: Record<string, number>;
    xlSheetVisible?: number;
    XlSheetVisibility?: Record<string, number>;
    wdExportFormatPDF?: number;
    WdExportFormat?: Record<string, number>;
  };
  ActiveDocument?: WpsDocument;
  ActiveWorkbook?: WpsWorkbook;
  ActiveSheet?: WpsWorksheet;
  Documents?: { Add(): WpsDocument };
  FileDialog?(type: number): WpsFileDialog;
  FileSystem?: WpsFileSystem;
  OAAssist?: { ShellExecute?(url: string, params?: string): unknown };
  Selection?: WpsSelection;
  WorksheetFunction?: { IsError(cell: WpsSheetCell): boolean | number };
  ActiveWindow?: {
    Hwnd?: number;
    SelectedSheets?: { Count: number; Item(index: number): WpsSheet };
  };
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
