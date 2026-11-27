import { describe, expect, it } from 'vitest';
import { MAX_DOCUMENT_PARAGRAPHS, MAX_DOCUMENT_TEXT } from '../../core/selection';
import type { WpsApplication, WpsDocument, WpsRange } from './bridge';
import {
  isWpsNormalStyle,
  nativeWpsStyle,
  readWpsDocumentStyles,
  supportsWpsDocumentStyles,
} from './document-styles';

interface Part {
  text: string;
  style: string | number;
  tableCount: number;
  listType: number;
  storyType: number;
  owner?: WpsDocument;
  duplicate?: (range: WpsRange) => WpsRange;
}

function setup(
  parts: Part[] = [{ text: '正文\r', style: -1, tableCount: 0, listType: 0, storyType: 1 }],
) {
  const document: WpsDocument = {
    Name: '文档.docx',
    FullName: 'C:\\docs\\文档.docx',
  };
  const offsets = () => {
    let start = 0;
    return parts.map((part) => {
      const current = { start, end: start + part.text.length };
      start = current.end;
      return current;
    });
  };
  const paragraphRange = (part: Part, bounds: { start: number; end: number }): WpsRange => {
    const range: WpsRange = {
      get Text() {
        return part.text;
      },
      get Start() {
        return bounds.start;
      },
      get End() {
        return bounds.end;
      },
      Font: { Size: 12 },
      get Document() {
        return part.owner ?? document;
      },
      get Duplicate() {
        return part.duplicate ? part.duplicate(range) : paragraphRange(part, bounds);
      },
      get StoryType() {
        return part.storyType;
      },
      get Style() {
        return part.style;
      },
      Tables: {
        get Count() {
          return part.tableCount;
        },
      },
      ListFormat: {
        get ListType() {
          return part.listType;
        },
      },
    };
    return range;
  };
  const content: WpsRange = {
    get Text() {
      return parts.map((part) => part.text).join('');
    },
    get Start() {
      return 0;
    },
    get End() {
      return parts.reduce((length, part) => length + part.text.length, 0);
    },
    Font: { Size: 12 },
    Document: document,
    get Duplicate() {
      return this;
    },
    StoryType: 1,
  };
  content.Paragraphs = {
    get Count() {
      return parts.length;
    },
    Item(index: number) {
      const part = parts[index - 1]!;
      const bounds = offsets()[index - 1]!;
      return { Range: paragraphRange(part, bounds) };
    },
  };
  document.Content = content;
  const app: WpsApplication = {
    ActiveDocument: document,
    Enum: {
      wdStyleNormal: -1,
      wdStyleHeading1: -2,
      wdStyleHeading2: -3,
      wdStyleHeading3: -4,
      wdListNoNumbering: 0,
      wdMainTextStory: 1,
    },
  };
  return { app, document, parts, content };
}

describe('WPS 全文样式读取边界', () => {
  it('使用实际运行时内置样式和列表/正文枚举，并按明确样式名归一化', () => {
    const fixture = setup([
      { text: '正文\r', style: 'Normal', tableCount: 0, listType: 0, storyType: 1 },
      { text: '标题\r', style: 'Heading 1', tableCount: 0, listType: 0, storyType: 1 },
      { text: '枚举标题\r', style: 'wdStyleHeading2', tableCount: 0, listType: 0, storyType: 1 },
      { text: '自定义\r', style: 'Nor_mal', tableCount: 0, listType: 0, storyType: 1 },
      { text: '表格标题\u0007', style: -1, tableCount: 0, listType: 2, storyType: 1 },
    ]);
    const state = readWpsDocumentStyles(fixture.app, fixture.document);
    expect(state.paragraphs.map(({ style }) => style)).toEqual([
      'Normal',
      'Heading1',
      'Heading2',
      'Other',
      'Normal',
    ]);
    expect(state.paragraphs.map(({ isListItem, inTable }) => [isListItem, inTable])).toEqual([
      [false, false],
      [false, false],
      [false, false],
      [false, false],
      [true, true],
    ]);
    expect(nativeWpsStyle(fixture.app, 'Heading2')).toBe(-3);
  });

  it('supports 允许分组枚举，但拒绝冲突、重复、缺失和非整数运行时值', () => {
    const fixture = setup();
    fixture.app.Enum = {
      WdBuiltinStyle: {
        wdStyleNormal: -1,
        wdStyleHeading1: -2,
        wdStyleHeading2: -3,
        wdStyleHeading3: -4,
      },
      WdListType: { wdListNoNumbering: 0 },
      WdStoryType: { wdMainTextStory: 1 },
    } as unknown as NonNullable<WpsApplication['Enum']>;
    expect(readWpsDocumentStyles(fixture.app, fixture.document).paragraphs).toHaveLength(1);

    fixture.app.Enum!.wdStyleHeading1 = -1;
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/枚举/u);
    fixture.app.Enum = {
      wdStyleNormal: 1.5,
      wdStyleHeading1: -2,
      wdStyleHeading2: -3,
      wdStyleHeading3: -4,
      wdListNoNumbering: 0,
      wdMainTextStory: 1,
    };
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/枚举/u);
    expect(() => nativeWpsStyle(fixture.app, 'Heading1')).toThrow(/枚举/u);
  });

  it('在写入前同时识别运行时枚举值和本地化 Normal 样式名', () => {
    const fixture = setup();
    expect(isWpsNormalStyle(fixture.app, -1)).toBe(true);
    expect(isWpsNormalStyle(fixture.app, 'Normal')).toBe(true);
    expect(isWpsNormalStyle(fixture.app, '正文')).toBe(true);
    expect(isWpsNormalStyle(fixture.app, 'Nor_mal')).toBe(false);
    expect(isWpsNormalStyle(fixture.app, -2)).toBe(false);
  });

  it.each([
    [
      'missing table count',
      (part: Part) => {
        (part as { tableCount?: number }).tableCount = undefined as unknown as number;
      },
    ],
    [
      'negative table count',
      (part: Part) => {
        part.tableCount = -1;
      },
    ],
    [
      'fractional table count',
      (part: Part) => {
        part.tableCount = 0.5;
      },
    ],
    [
      'missing list type',
      (part: Part) => {
        (part as { listType?: number }).listType = undefined as unknown as number;
      },
    ],
    [
      'negative list type',
      (part: Part) => {
        part.listType = -1;
      },
    ],
    [
      'unknown story',
      (part: Part) => {
        part.storyType = 99;
      },
    ],
    [
      'missing style',
      (part: Part) => {
        (part as { style?: string | number }).style = undefined;
      },
    ],
  ] as const)('对 %s 拒绝全文读取', (_name, corrupt) => {
    const fixture = setup();
    corrupt(fixture.parts[0]!);
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow();
    expect(supportsWpsDocumentStyles(fixture.app)).toBe(false);
  });

  it('要求当前文档和每个段落都有可验证归属，姓名相同不能冒充同一文档', () => {
    const fixture = setup();
    fixture.app.ActiveDocument = { ...fixture.document, FullName: 'C:\\other\\文档.docx' };
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/不属于当前文档/u);
    fixture.app.ActiveDocument = fixture.document;

    fixture.parts[0]!.owner = {
      Name: fixture.document.Name,
      FullName: 'C:\\other\\文档.docx',
    };
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/不属于当前文档/u);
    fixture.parts[0]!.owner = { Name: fixture.document.Name };
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/不属于当前文档/u);
  });

  it('要求 Duplicate 为独立且文本、位置、样式、列表和表格状态一致的范围', () => {
    const fixture = setup();
    fixture.parts[0]!.duplicate = (range) => range;
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/独立/u);

    fixture.parts[0]!.duplicate = (range) => {
      const duplicate = Object.create(range) as WpsRange;
      Object.defineProperty(duplicate, 'Text', { value: '被换掉的文本\r' });
      return duplicate;
    };
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/范围已改变/u);
  });

  it('段落数和总文本长度在读取前拒绝，达到上限则仍允许', () => {
    const tooMany = setup(
      Array.from({ length: MAX_DOCUMENT_PARAGRAPHS + 1 }, () => ({
        text: '\r',
        style: -1,
        tableCount: 0,
        listType: 0,
        storyType: 1,
      })),
    );
    expect(() => readWpsDocumentStyles(tooMany.app, tooMany.document)).toThrow(/段落超出/u);

    const tooLong = setup([
      {
        text: '长'.repeat(MAX_DOCUMENT_TEXT + 1),
        style: -1,
        tableCount: 0,
        listType: 0,
        storyType: 1,
      },
    ]);
    expect(() => readWpsDocumentStyles(tooLong.app, tooLong.document)).toThrow(/字符数超出/u);
    expect(supportsWpsDocumentStyles(tooLong.app)).toBe(false);

    const atLimit = setup([
      { text: '长'.repeat(MAX_DOCUMENT_TEXT), style: -1, tableCount: 0, listType: 0, storyType: 1 },
    ]);
    expect(readWpsDocumentStyles(atLimit.app, atLimit.document).paragraphs[0]!.text).toHaveLength(
      MAX_DOCUMENT_TEXT,
    );
  });

  it('对不连续和正文以外的段落范围失败关闭', () => {
    const fixture = setup([
      { text: '第一段\r', style: -1, tableCount: 0, listType: 0, storyType: 1 },
      { text: '第二段\r', style: -1, tableCount: 0, listType: 0, storyType: 1 },
    ]);
    fixture.parts[1]!.duplicate = (range) => {
      const duplicate = Object.create(range) as WpsRange;
      Object.defineProperty(duplicate, 'Start', { value: range.Start + 1 });
      return duplicate;
    };
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/范围已改变/u);

    fixture.parts[1]!.duplicate = undefined;
    fixture.parts[1]!.storyType = 2;
    expect(() => readWpsDocumentStyles(fixture.app, fixture.document)).toThrow(/正文范围/u);
  });

  it('正向能力探测只证明当前文档的同步只读接口可用', () => {
    const fixture = setup();
    expect(supportsWpsDocumentStyles(fixture.app)).toBe(true);
    expect(fixture.parts[0]!.style).toBe(-1);
    expect(fixture.app.ActiveDocument).toBe(fixture.document);
  });
});
