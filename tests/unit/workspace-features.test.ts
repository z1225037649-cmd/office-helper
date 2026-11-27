import { describe, expect, it } from 'vitest';
import type {
  CellType,
  CellValue,
  DuplicateOptions,
  NoticeFields,
  PrintSettingsSnapshot,
  SelectionSnapshot,
  SheetLayoutSnapshot,
} from '../../src/core/contracts';
import { noticeText } from '../../src/core/extended-actions';
import { checkRecords } from '../../src/features/check-records/feature';
import { noticeTemplate } from '../../src/features/notice-template/feature';
import { printCheck } from '../../src/features/print-check/feature';
import { exportPdf, saveCopy } from '../../src/features/save-copy/feature';
import {
  autofitColumns,
  emphasizeHeader,
  wrapCells,
} from '../../src/features/sheet-layout/feature';

function valueType(value: CellValue): CellType {
  if (value === null || value === '') return 'empty';
  if (typeof value === 'string') return 'text';
  if (typeof value === 'number') return 'number';
  return 'boolean';
}

function sheet(
  values: CellValue[][] = [
    ['姓名', '电话'],
    ['张三', '123'],
    ['李四', '456'],
  ],
): SelectionSnapshot {
  const columns = values[0]?.length ?? 0;
  return {
    documentId: 'sheet-id',
    documentName: '记录.xlsx',
    kind: 'sheet',
    targetId: 'Sheet1!D5:E7',
    targetLabel: 'D5:E7',
    revision: '1',
    values,
    cellTypes: values.map((row) => row.map(valueType)),
    rowScope: 'all',
    sheetLayout: {
      columnWidths: Array.from({ length: columns }, (_, c) => 68 + c * 24),
      widthUnit: 'points',
      wrapText: values.map((row) => row.map(() => false)),
      headerBold: Array.from({ length: columns }, () => false),
      merged: false,
      protected: false,
      firstRow: 4,
      firstColumn: 3,
      sheetName: '登记表',
    },
    printSettings: {
      available: true,
      printArea: 'A1:B9',
      orientation: 'Portrait',
      zoom: 80,
      fitWide: 2,
      fitTall: 3,
      contentWidthPoints: 1600,
    },
  };
}

function text(): SelectionSnapshot {
  return { ...sheet(), kind: 'text', documentName: '原稿.docx', text: '' };
}

function duplicates(overrides: Partial<DuplicateOptions> = {}): DuplicateOptions {
  return {
    keyColumns: [0],
    requiredColumns: [],
    trimWhitespace: false,
    ignoreCase: false,
    hasHeader: false,
    ...overrides,
  };
}

function recordReport(values: CellValue[][], overrides: Partial<DuplicateOptions> = {}) {
  const { plan } = checkRecords.prepare(sheet(values), { duplicates: duplicates(overrides) });
  if (plan.type !== 'report') throw new Error('检查必须只读');
  return plan;
}

const notice: NoticeFields = {
  title: '活动通知',
  time: '2026年11月28日 09:00',
  location: '会议室',
  body: '请准时参加。\n请自带水杯。',
  contact: '办公室 12345',
};

const layoutFeatures = [autofitColumns, wrapCells, emphasizeHeader];
const allFeatures = [
  ...layoutFeatures,
  printCheck,
  saveCopy,
  exportPdf,
  checkRecords,
  noticeTemplate,
];

describe('新增工作区功能的独立定义', () => {
  it('八个导出具有唯一约定 ID，均提供必要文案与只读选区能力', () => {
    expect(allFeatures.map((feature) => feature.id)).toEqual([
      'sheet.autofit',
      'sheet.wrap',
      'sheet.header',
      'sheet.print',
      'document.saveCopy',
      'document.exportPdf',
      'sheet.checkRecords',
      'document.notice',
    ]);
    for (const feature of allFeatures) {
      expect(feature.title).not.toBe('');
      expect(feature.description).not.toBe('');
      expect(feature.phrases.length).toBeGreaterThan(0);
      expect(feature.requires).toContain('selection.read');
    }
  });

  it.each(layoutFeatures)('$id 仅用于表格，并单独申请完整布局与对应写能力', (feature) => {
    expect(feature.kinds).toEqual(['sheet']);
    expect(feature.requires).toEqual(['selection.read', feature.id]);
    expect(feature.readOptions).toEqual({ includeSheetLayout: true });
  });

  it('打印保留无原生写能力时的指引入口，判重显式申请每格类型', () => {
    expect(printCheck.requires).toEqual(['selection.read']);
    expect(printCheck.readOptions).toEqual({ includePrintSettings: true });
    expect(checkRecords.requires).toEqual(['selection.read']);
    expect(checkRecords.readOptions).toEqual({ includeCellTypes: true });
    expect(checkRecords.form).toBe('duplicates');
    expect(checkRecords.kinds).toEqual(['sheet']);
  });

  it.each([saveCopy, exportPdf])('$id 两种文件均进入保存表单', (feature) => {
    expect(feature.kinds).toEqual(['text', 'sheet']);
    expect(feature.form).toBe('save');
    expect(feature.requires).toContain(feature.id);
  });

  it('通知进入文字文档专用表单', () => {
    expect(noticeTemplate.kinds).toEqual(['text']);
    expect(noticeTemplate.form).toBe('notice');
    expect(noticeTemplate.requires).toContain('document.notice');
  });

  it.each(['找出重复和漏填', '找重复和漏填', '找出重复', '找出漏填'])(
    '包含用户查重原话 %s',
    (phrase) => {
      expect(checkRecords.phrases).toContain(phrase);
    },
  );

  it.each(['保存成PDF', '保存成 PDF'])('包含用户 PDF 原话 %s', (phrase) => {
    expect(exportPdf.phrases).toContain(phrase);
  });
});

describe('表格格式的预览与可靠性边界', () => {
  it('列宽展示原单位原数值，说明整列影响且不编造目标宽度', () => {
    const selection = sheet();
    const before = structuredClone(selection);
    const result = autofitColumns.prepare(selection);
    expect(result.plan).toEqual({ type: 'sheet.autofit' });
    expect(result.summary).toContain('第 4 列 68 磅');
    expect(result.summary).toContain('第 5 列 92 磅');
    expect(result.summary).toContain('整列');
    expect(result.summary).toContain('选区外');
    expect(result.summary).toContain('原生表格按内容自动计算');
    expect(selection).toEqual(before);
  });

  it('不将字符单位误标成磅，允许原始零列宽', () => {
    const selection = sheet();
    selection.sheetLayout!.widthUnit = 'characters';
    selection.sheetLayout!.columnWidths = [0, 12.75];
    expect(autofitColumns.prepare(selection).summary).toContain('12.75 字符宽度单位');
  });

  it('换行预览保留逐格已知差异，只针对所选单元格', () => {
    const selection = sheet();
    selection.sheetLayout!.wrapText[1]![0] = true;
    const before = structuredClone(selection);
    const result = wrapCells.prepare(selection);
    expect(result.plan).toEqual({ type: 'sheet.wrap' });
    expect(result.summary).toContain('6 个单元格');
    expect(result.summary).toContain('1 格已启用、5 格未启用');
    expect(result.summary).toContain('目标：全部启用');
    expect(selection).toEqual(before);
  });

  it('表头仅处理偏移选区第一行，并显示每列原粗细', () => {
    const selection = sheet();
    selection.sheetLayout!.headerBold = [false, true];
    const before = structuredClone(selection);
    const result = emphasizeHeader.prepare(selection);
    expect(result.plan).toEqual({ type: 'sheet.header' });
    expect(result.summary).toContain('选区第一行（工作表第 5 行）');
    expect(result.summary).toContain('第 1 列未加粗；选区第 2 列已加粗');
    expect(result.summary).toContain('不处理选区后续行');
    expect(selection).toEqual(before);
  });

  it('已全部启用换行或加粗时不发出重复写入计划', () => {
    const selection = sheet();
    selection.sheetLayout!.wrapText = selection.values!.map((row) => row.map(() => true));
    selection.sheetLayout!.headerBold = [true, true];
    expect(() => wrapCells.prepare(selection)).toThrow('无需修改');
    expect(() => emphasizeHeader.prepare(selection)).toThrow('无需修改');
  });

  it.each(layoutFeatures)('$id 拒绝文字、空选区和非矩形选区', (feature) => {
    expect(() => feature.prepare(text())).toThrow('表格');
    expect(() => feature.prepare(sheet([]))).toThrow('请先选择');
    expect(() => feature.prepare(sheet([[]]))).toThrow('请先选择');
    expect(() => feature.prepare(sheet([[1, 2], [3]]))).toThrow('连续');
    const noLayout = sheet();
    delete noLayout.sheetLayout;
    expect(() => feature.prepare(noLayout)).toThrow('完整布局');
  });

  const badLayouts: [string, object][] = [
    ['合并', { merged: true }],
    ['保护', { protected: true }],
    ['未知合并状态', { merged: undefined }],
    ['未知保护状态', { protected: null }],
    ['负行号', { firstRow: -1 }],
    ['小数列号', { firstColumn: 0.5 }],
    ['缺工作表名', { sheetName: ' ' }],
    ['未知单位', { widthUnit: 'pixels' }],
    ['缺列宽', { columnWidths: [12] }],
    ['非法列宽', { columnWidths: [NaN, 12] }],
    ['负列宽', { columnWidths: [-1, 12] }],
    ['字符串列宽', { columnWidths: ['12', 12] }],
    ['缺换行行', { wrapText: [[false, false]] }],
    ['缺换行列', { wrapText: [[false], [false], [false]] }],
    [
      '未知换行状态',
      {
        wrapText: [
          [null, false],
          [false, false],
          [false, false],
        ],
      },
    ],
    ['缺粗细', { headerBold: [false] }],
    ['未知粗细', { headerBold: [null, false] }],
    ['稀疏粗细', { headerBold: new Array(2) }],
  ];
  for (const feature of layoutFeatures) {
    it.each(badLayouts)(`${feature.id} 拒绝不可靠布局：%s`, (_label, patch) => {
      const selection = sheet();
      selection.sheetLayout = { ...selection.sheetLayout!, ...patch } as SheetLayoutSnapshot;
      expect(() => feature.prepare(selection)).toThrow();
    });
  }
});

describe('打印前检查', () => {
  it('预览原值和目标参数，提示实际分页与文字过小风险', () => {
    const selection = sheet();
    const before = structuredClone(selection);
    const result = printCheck.prepare(selection);
    expect(result.plan).toEqual({ type: 'sheet.print' });
    for (const detail of [
      'A1:B9',
      '纵向',
      '80%',
      '宽 2 页、高 3 页',
      '当前选区',
      '横向',
      '宽度 1 页',
      '高度自动',
      '实际页数需在原生打印预览中核对',
      '文字过小',
      '不会提交打印',
    ]) {
      expect(result.summary).toContain(detail);
    }
    expect(result.summary).not.toMatch(/预计\s*\d+\s*页|共\s*\d+\s*页/u);
    expect(result.confirmLabel).toBeTruthy();
    expect(selection).toEqual(before);
  });

  it('可读出原有按页适配、自动页宽页高和未指定打印区域', () => {
    const selection = sheet();
    Object.assign(selection.printSettings!, { printArea: '', zoom: false, fitWide: 0, fitTall: 0 });
    const { summary } = printCheck.prepare(selection);
    expect(summary).toContain('未指定');
    expect(summary).toContain('缩放 按页适配；宽 自动、高 自动');
  });

  it.each([false, undefined])('打印接口不可用 %s 时只给原生指引，不生成写计划', (available) => {
    const selection = sheet();
    if (available === false)
      selection.printSettings = {
        available: false,
        reason: '需要更高版本。',
      } as PrintSettingsSnapshot;
    else delete selection.printSettings;
    const result = printCheck.prepare(selection);
    expect(result.plan.type).toBe('report');
    if (result.plan.type !== 'report') throw new Error('必须只读');
    expect(result.plan.message).toContain('文件 → 打印');
    expect(result.plan.message).toContain('本次未修改');
    expect(result.plan.message).toContain('原生打印预览');
    expect(result.plan.message).toContain('不会提交打印');
  });

  const invalidPrint: [string, object][] = [
    ['未知可用状态', { available: 1 }],
    ['缺区域原值', { printArea: undefined }],
    ['空方向', { orientation: '' }],
    ['非法缩放', { zoom: NaN }],
    ['零缩放', { zoom: 0 }],
    ['负宽页数', { fitWide: -1 }],
    ['小数高页数', { fitTall: 0.5 }],
    ['未知内容宽度', { contentWidthPoints: Infinity }],
  ];
  it.each(invalidPrint)('拒绝不可靠打印原值：%s', (_label, patch) => {
    const selection = sheet();
    selection.printSettings = { ...selection.printSettings!, ...patch } as PrintSettingsSnapshot;
    expect(() => printCheck.prepare(selection)).toThrow('原值读取不完整');
  });

  it('只在表格的有效连续选区上提供设置计划', () => {
    expect(() => printCheck.prepare(text())).toThrow('只支持表格');
    expect(() => printCheck.prepare(sheet([]))).toThrow('矩形区域');
    expect(() => printCheck.prepare(sheet([[1], [2, 3]]))).toThrow('矩形区域');
  });
});

describe('另存副本和导出 PDF', () => {
  it.each(['selection', 'activeSheet'] as const)(
    'PDF 显式选择 %s，并在预览列明实际范围',
    (exportScope) => {
      const prepared = exportPdf.prepare(sheet(), { fileName: '范围.pdf', exportScope });
      expect(prepared.plan).toMatchObject({ type: 'document.exportPdf', scope: exportScope });
      expect(prepared.summary).toContain(
        exportScope === 'selection' ? '仅当前选区' : '仅当前工作表',
      );
      expect(prepared.summary).not.toContain('整个工作簿（所有工作表）');
    },
  );

  it('拒绝空文字选区、文字的活动表范围及部分副本', () => {
    expect(() =>
      exportPdf.prepare(text(), { fileName: '范围.pdf', exportScope: 'selection' }),
    ).toThrow('选中');
    expect(() =>
      exportPdf.prepare(text(), { fileName: '范围.pdf', exportScope: 'activeSheet' }),
    ).toThrow('表格');
    expect(() =>
      saveCopy.prepare(sheet(), { fileName: '副本.xlsx', exportScope: 'selection' }),
    ).toThrow('整份原稿');
  });

  it('非法范围被核心校验拒绝，文字选区可带入合法计划', () => {
    expect(() =>
      exportPdf.prepare(sheet(), { fileName: '范围.pdf', exportScope: 'all' as never }),
    ).toThrow('范围');
    expect(
      exportPdf.prepare(
        { ...text(), text: '所选内容' },
        { fileName: '范围.pdf', exportScope: 'selection' },
      ).plan,
    ).toMatchObject({ scope: 'selection' });
  });

  for (const feature of [saveCopy, exportPdf]) {
    it.each([undefined, {}, { fileName: '' }, { fileName: '  ' }])(
      `${feature.id} 无明确文件名时拒绝`,
      (parameters) => {
        expect(() => feature.prepare(text(), parameters)).toThrow('文件名');
      },
    );
    it.each(['../副本.pdf', 'C:\\副本.pdf', 'CON.pdf', '副本?.pdf', '副本.pdf ', '副本.'])(
      `${feature.id} 复用公共文件名校验拒绝 %s`,
      (fileName) => {
        expect(() => feature.prepare(text(), { fileName })).toThrow();
      },
    );
    it.each(['folder\npath', 'x'.repeat(1001)])(`${feature.id} 拒绝无效文件夹`, (folder) => {
      const fileName = feature === saveCopy ? '新副本.docx' : '通知.pdf';
      expect(() => feature.prepare(text(), { fileName, folder })).toThrow('文件夹');
    });
    it(`${feature.id} 名称比较忽略大小写，指定不同文件夹也不能同名`, () => {
      const selection = text();
      selection.documentName = feature === saveCopy ? 'C:\\原件\\原稿.docx' : 'C:/原件/原稿.pdf';
      const fileName = feature === saveCopy ? '原稿.DOCX' : '原稿.PDF';
      expect(() => feature.prepare(selection, { fileName, folder: 'D:\\副本' })).toThrow(
        '不同于原稿',
      );
    });
  }

  it.each([text(), sheet()])(
    '另存范围为整份文件，提供确认文案且不把预览说成已保存',
    (selection) => {
      const fileName = selection.kind === 'text' ? '副本.docx' : '副本.xlsx';
      const parameters = { fileName, folder: 'D:\\备份' };
      const before = structuredClone({ selection, parameters });
      const result = saveCopy.prepare(selection, parameters);
      expect(result.plan).toEqual({ type: 'document.saveCopy', ...parameters });
      expect(result.confirmLabel).toBe('确认另存');
      expect(result.summary).toContain(
        selection.kind === 'text' ? '整份文档' : '整个工作簿（所有工作表）',
      );
      expect(result.summary).toContain('不以当前选区');
      expect(result.summary).toContain('保留原件');
      expect(result.summary).toContain(parameters.folder);
      expect(result.summary).toContain('重名');
      expect(result.summary).toContain('操作预览');
      expect(result.summary).not.toContain('已保存');
      expect({ selection, parameters }).toEqual(before);
    },
  );

  it.each([text(), sheet()])('PDF 范围为整份文件或所有工作表，并提供单独确认', (selection) => {
    const result = exportPdf.prepare(selection, { fileName: '导出.PDF' });
    expect(result.plan).toEqual({ type: 'document.exportPdf', fileName: '导出.PDF' });
    expect(result.confirmLabel).toBe('确认导出');
    expect(result.summary).toContain(
      selection.kind === 'text' ? '整份文档' : '整个工作簿（所有工作表）',
    );
    expect(result.summary).toContain('原生保存对话框');
    expect(result.summary).toContain('取消对话框不代表导出成功');
    expect(result.summary).not.toContain('已导出');
  });

  it.each(['副本.docx', '副本.xlsx', '副本.pdf.exe', '副本'])(
    'PDF 拒绝非 PDF 后缀 %s',
    (fileName) => {
      expect(() => exportPdf.prepare(text(), { fileName })).toThrow('.pdf');
    },
  );

  it.each(['.xlsm', '副本.xlsx', '副本.xls', '副本.pdf', '副本.xlsm.exe', '副本'])(
    '宏工作簿不允许伪装副本后缀：%s',
    (fileName) => {
      const selection = sheet();
      selection.documentName = 'C:\\资料\\原稿.xlsm';
      expect(() => saveCopy.prepare(selection, { fileName })).toThrow('原稿后缀 .xlsm');
    },
  );

  it.each(['原稿.xlsm', '原稿.XLSM'])('工作簿同名被拒绝：%s', (fileName) => {
    const selection = sheet();
    selection.documentName = 'C:\\资料\\原稿.xlsm';
    expect(() => saveCopy.prepare(selection, { fileName })).toThrow('同名覆盖');
  });

  it.each(['原稿.xlsm', '原稿.XLSM', '原稿.docm', '原稿.xls', '原稿.doc'])(
    '原后缀保留，比较不区分大小写：%s',
    (documentName) => {
      const selection = sheet();
      selection.documentName = `C:/资料/${documentName}`;
      const extension = documentName.slice(documentName.lastIndexOf('.'));
      const fileName = `新副本${extension.toUpperCase()}`;
      expect(saveCopy.prepare(selection, { fileName }).plan).toEqual({
        type: 'document.saveCopy',
        fileName,
      });
    },
  );

  it.each(['text', 'sheet'] as const)('无后缀 %s 使用对应 docx/xlsx 后缀', (kind) => {
    const selection = kind === 'text' ? text() : sheet();
    selection.documentName = '未命名';
    const extension = kind === 'text' ? '.docx' : '.xlsx';
    expect(saveCopy.prepare(selection, { fileName: `副本${extension}` }).plan.type).toBe(
      'document.saveCopy',
    );
    expect(() => saveCopy.prepare(selection, { fileName: '副本' })).toThrow(extension);
    expect(() => saveCopy.prepare(selection, { fileName: '副本.pdf' })).toThrow(extension);
  });
});

describe('记录判重与漏填的纯统计', () => {
  it('数字、文字、布尔值保留类型，不因字符串转换碰撞', () => {
    const result = recordReport([[1], ['1'], [true], ['true'], [1], ['1'], [true], ['true']]);
    expect(result.message).toContain('疑似重复 4 行');
    expect(result.message).toContain('共 4 组');
    expect(result.findings?.map(({ row, column }) => [row, column])).toEqual([
      [4, 0],
      [5, 0],
      [6, 0],
      [7, 0],
    ]);
    expect(result.message).toContain('数字 1 与文字“1”不同');
  });

  it('复合键同时判断姓名和电话，不用有碰撞风险的字符串分隔符', () => {
    const result = recordReport(
      [
        ['张三', '111'],
        ['张三', '222'],
        ['张三', '111'],
        ['a|b', 'c'],
        ['a', 'b|c'],
        ['a', 'b|c'],
        ['x","y', 'z'],
        ['x', 'y","z'],
      ],
      { keyColumns: [0, 1] },
    );
    expect(result.message).toContain('疑似重复 2 行');
    expect(result.findings?.map(({ row }) => row)).toEqual([2, 5]);
    expect(result.findings?.[0]?.description).toContain('选区第 1 行');
  });

  it.each([
    [false, false, 0],
    [true, false, 1],
    [false, true, 1],
    [true, true, 2],
  ])('去首尾空格=%s、忽略大小写=%s 的规则可解释', (trimWhitespace, ignoreCase, expected) => {
    const result = recordReport([[' Alice '], ['Alice'], ['alice']], {
      trimWhitespace: Boolean(trimWhitespace),
      ignoreCase: Boolean(ignoreCase),
    });
    expect(result.message).toContain(`疑似重复 ${expected} 行`);
    expect(result.message).toContain(`${trimWhitespace ? '去除' : '保留'}首尾空格`);
    expect(result.message).toContain(`${ignoreCase ? '忽略' : '区分'}大小写`);
  });

  it('只比较首尾空格和大小写，不删除中间空白', () => {
    const result = recordReport([['A B'], ['AB'], ['a b']], {
      trimWhitespace: true,
      ignoreCase: true,
    });
    expect(result.message).toContain('疑似重复 1 行');
    expect(result.findings?.[0]?.row).toBe(2);
  });

  it('hasHeader 排除首行但 findings 保留原始选区零基坐标', () => {
    const result = recordReport(
      [
        ['姓名', '电话'],
        ['张三', '1'],
        ['张三', null],
      ],
      {
        keyColumns: [0],
        requiredColumns: [1],
        hasHeader: true,
      },
    );
    expect(result.message).toContain('已扫描 2 行');
    expect(result.findings?.map(({ row, column }) => [row, column])).toEqual([
      [2, 0],
      [2, 1],
    ]);
    expect(result.findings?.[0]?.description).toContain('选区第 2 行');
    expect(result.message).toContain('首行为表头，不参与检查');
  });

  it('首行可以是数据；同一值三行只计两行重复，首次行号 0 不丢失', () => {
    const result = recordReport([['x'], ['x'], ['x']]);
    expect(result.message).toContain('疑似重复 2 行');
    expect(result.message).toContain('共 1 组');
    expect(result.findings?.every((finding) => finding.description.includes('选区第 1 行'))).toBe(
      true,
    );
  });

  it('空值、空串、全空白不互相判重；0 和 false 是有效必填值', () => {
    const result = recordReport([[null], [''], [' \t '], [0], [false], [0], [false]], {
      requiredColumns: [0],
    });
    expect(result.message).toContain('疑似重复 2 行');
    expect(result.message).toContain('疑似漏填 3 格（涉及 3 行）');
    expect(result.message).toContain('因判断列空白跳过判重 3 行');
    expect(
      result.findings?.filter((finding) => finding.id.startsWith('missing')).map(({ row }) => row),
    ).toEqual([0, 1, 2]);
  });

  it('复合键任一列为空就跳过判重，但仍检查必填项', () => {
    const result = recordReport(
      [
        ['张三', null],
        ['张三', ''],
        ['张三', '1'],
      ],
      {
        keyColumns: [0, 1],
        requiredColumns: [1],
      },
    );
    expect(result.message).toContain('疑似重复 0 行');
    expect(result.message).toContain('疑似漏填 2 格');
    expect(result.message).toContain('跳过判重 2 行');
  });

  it('只查漏填时按格计数、另报涉及行数，列索引相对选区', () => {
    const result = recordReport(
      [
        ['a', '', null],
        ['b', 'x', null],
      ],
      {
        keyColumns: [],
        requiredColumns: [1, 2],
      },
    );
    expect(result.message).toContain('疑似重复 0 行');
    expect(result.message).toContain('疑似漏填 3 格（涉及 2 行）');
    expect(result.findings?.map(({ row, column }) => [row, column])).toEqual([
      [0, 1],
      [0, 2],
      [1, 2],
    ]);
  });

  it('只有表头或没有问题时可返回零条的只读报告', () => {
    const headerOnly = recordReport([['姓名']], { hasHeader: true });
    expect(headerOnly.message).toContain('已扫描 0 行');
    expect(headerOnly.findings).toEqual([]);
    expect(recordReport([['a'], ['b']]).findings).toEqual([]);
  });

  it('展示截断到 200 条后继续完整计数，报告 ID 唯一且坐标不越界', () => {
    const values: CellValue[][] = Array.from({ length: 1000 }, () => ['重复', null]);
    const selection = sheet(values);
    const options = duplicates({ requiredColumns: [1] });
    const before = structuredClone({ selection, options });
    const { plan } = checkRecords.prepare(selection, { duplicates: options });
    if (plan.type !== 'report') throw new Error('必须只读');
    expect(plan.message).toContain('疑似重复 999 行');
    expect(plan.message).toContain('疑似漏填 1000 格（涉及 1000 行）');
    expect(plan.message).toContain('共 1999 条结果，仅显示前 200 条');
    expect(plan.message).toContain('覆盖全部 1000 行');
    expect(plan.findings).toHaveLength(200);
    expect(new Set(plan.findings!.map(({ id }) => id)).size).toBe(200);
    expect(
      plan.findings!.every(
        ({ row, column }) => row >= 0 && row < 1000 && column >= 0 && column < 2,
      ),
    ).toBe(true);
    expect({ selection, options }).toEqual(before);
  });

  it('恰好 200 条无需声称截断', () => {
    const result = recordReport(Array.from({ length: 201 }, () => ['x']));
    expect(result.findings).toHaveLength(200);
    expect(result.message).toContain('已全部显示');
    expect(result.message).not.toContain('仅显示前');
  });

  it('不同错误码表现的错误单元格整行排除，不与同码普通文字或数字混淆', () => {
    const selection = sheet([
      ['#N/A', null],
      ['#N/A', null],
      [2042, null],
      [2042, null],
      ['#N/A', 'x'],
      ['#N/A', 'y'],
    ]);
    selection.cellTypes![0]![0] = 'error';
    selection.cellTypes![1]![0] = 'error';
    selection.cellTypes![2]![0] = 'error';
    selection.cellTypes![3]![0] = 'error';
    const { plan } = checkRecords.prepare(selection, {
      duplicates: duplicates({ requiredColumns: [1] }),
    });
    if (plan.type !== 'report') throw new Error('必须只读');
    expect(plan.message).toContain('4 行含 4 个错误单元格，整行排除，实际检查 2 行');
    expect(plan.message).toContain('疑似重复 1 行');
    expect(plan.message).toContain('疑似漏填 0 格');
    expect(plan.findings?.map(({ row }) => row)).toEqual([5]);
  });

  it('错误出现在非判断列时也整行排除，并明确报告', () => {
    const selection = sheet([
      ['张三', '#REF!'],
      ['张三', '#REF!'],
    ]);
    selection.cellTypes!.forEach((row) => {
      row[1] = 'error';
    });
    const { plan } = checkRecords.prepare(selection, { duplicates: duplicates() });
    if (plan.type !== 'report') throw new Error('必须只读');
    expect(plan.message).toContain('2 行含 2 个错误单元格');
    expect(plan.message).toContain('疑似重复 0 行');
    expect(plan.findings).toEqual([]);
  });

  const badTypeSnapshots: [string, object][] = [
    ['缺少类型', { cellTypes: undefined }],
    ['类型矩阵缺行', { cellTypes: [['text', 'text']] }],
    [
      '未知类型',
      {
        cellTypes: [
          ['unknown', 'text'],
          ['text', 'text'],
          ['text', 'text'],
        ],
      },
    ],
    ['类型矩阵缺列', { cellTypes: [['text'], ['text'], ['text']] }],
    [
      '类型与值不符',
      {
        cellTypes: [
          ['number', 'text'],
          ['text', 'text'],
          ['text', 'text'],
        ],
      },
    ],
    ['完整行范围未知', { rowScope: undefined }],
  ];
  it.each(badTypeSnapshots)('不可靠类型与范围拒绝：%s', (_label, patch) => {
    const selection = { ...sheet(), ...patch } as SelectionSnapshot;
    expect(() => checkRecords.prepare(selection, { duplicates: duplicates() })).toThrow('可靠读取');
  });

  const badOptions: [string, object][] = [
    ['未选列', { keyColumns: [], requiredColumns: [] }],
    ['负列', { keyColumns: [-1] }],
    ['越界列', { keyColumns: [2] }],
    ['小数列', { keyColumns: [0.5] }],
    ['字符串列', { keyColumns: ['0'] }],
    ['重复判断列', { keyColumns: [0, 0] }],
    ['重复必填列', { requiredColumns: [1, 1] }],
    ['越界必填列', { requiredColumns: [2] }],
    ['缺少列数组', { keyColumns: undefined }],
    ['未知表头选项', { hasHeader: undefined }],
    ['数字大小写选项', { ignoreCase: 1 }],
    ['字符串空格选项', { trimWhitespace: 'true' }],
  ];
  it.each(badOptions)('拒绝非法选项：%s', (_label, patch) => {
    expect(() =>
      checkRecords.prepare(sheet(), {
        duplicates: { ...duplicates(), ...patch } as DuplicateOptions,
      }),
    ).toThrow();
  });

  it('缺参数、空数据、非矩形、非有限数字或错误文件类型都明确拒绝', () => {
    expect(() => checkRecords.prepare(sheet())).toThrow('请选择');
    expect(() => checkRecords.prepare(text(), { duplicates: duplicates() })).toThrow('表格');
    for (const values of [[], [[]], [[1, 2], [3]], [[NaN]], [[Infinity]]]) {
      expect(() => checkRecords.prepare(sheet(values), { duplicates: duplicates() })).toThrow();
    }
  });
});

describe('通知模板', () => {
  it('五字段先去首尾空白，完整预览正文；计划独立复制输入且不要求选中文字', () => {
    const fields = Object.fromEntries(
      Object.entries(notice).map(([key, value]) => [key, `  ${value}  `]),
    ) as unknown as NoticeFields;
    const selection = text();
    const before = structuredClone({ selection, fields });
    const result = noticeTemplate.prepare(selection, { notice: fields });
    expect(result.plan).toEqual({ type: 'document.notice', fields: notice });
    expect(result.summary).toContain(noticeText(notice));
    expect(result.summary).toContain('新文档');
    expect(result.summary).toContain('不覆盖当前文档');
    expect(result.confirmLabel).toBe('确认生成通知');
    expect({ selection, fields }).toEqual(before);
    fields.title = '之后的改动';
    if (result.plan.type !== 'document.notice') throw new Error('错误计划');
    expect(result.plan.fields.title).toBe(notice.title);
  });

  it('缺失表单或错误类型均明确拒绝', () => {
    expect(() => noticeTemplate.prepare(text())).toThrow('请填写');
    expect(() => noticeTemplate.prepare(sheet(), { notice })).toThrow('文字文档');
  });

  const limits: [keyof NoticeFields, number][] = [
    ['title', 120],
    ['time', 200],
    ['location', 300],
    ['body', 12000],
    ['contact', 300],
  ];
  it.each(limits)('%s 必填且必须是字符串', (field) => {
    for (const value of ['', ' \t ', undefined, 123]) {
      const fields = { ...notice, [field]: value } as NoticeFields;
      expect(() => noticeTemplate.prepare(text(), { notice: fields })).toThrow('请填写');
    }
  });

  it.each(limits)('%s 接受去空格后限长 %d，超过一字拒绝', (field, limit) => {
    const fields = { ...notice, [field]: `  ${'字'.repeat(limit)}  ` };
    expect(noticeTemplate.prepare(text(), { notice: fields }).plan.type).toBe('document.notice');
    fields[field] = '字'.repeat(limit + 1);
    expect(() => noticeTemplate.prepare(text(), { notice: fields })).toThrow('缩短');
  });

  it.each(['title', 'time', 'location', 'contact'] as const)('%s 不允许中间或边缘换行', (field) => {
    for (const value of ['前\n后', '前\r后', '\n前', '前\n', '前\u2028后']) {
      expect(() =>
        noticeTemplate.prepare(text(), { notice: { ...notice, [field]: value } }),
      ).toThrow('各填写一行');
    }
  });

  it('正文允许多段和换行，禁止非法控制字符', () => {
    const fields = { ...notice, body: '第一段\r\n\r\n第二段\n第三段' };
    expect(noticeTemplate.prepare(text(), { notice: fields }).summary).toContain(fields.body);
    fields.body = '正文\u0000尾部';
    expect(() => noticeTemplate.prepare(text(), { notice: fields })).toThrow();
  });
});
