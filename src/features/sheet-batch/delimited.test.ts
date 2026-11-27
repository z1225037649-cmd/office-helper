import { describe, expect, it } from 'vitest';
import {
  detectDelimiter,
  formulaLikeTextCells,
  parseDelimitedFile,
  parseDelimitedText,
  serializeSheetBatchCsv,
} from './delimited';

describe('CSV / TSV 批量导入与导出', () => {
  it('自动识别逗号和制表符，保留字符串、空字段与行顺序', () => {
    expect(parseDelimitedText('姓名,电话\r\n张三,123\r\n李四,', '名单.csv')).toEqual({
      sourceLabel: '名单.csv',
      headers: ['姓名', '电话'],
      rows: [
        ['张三', '123'],
        ['李四', ''],
      ],
    });
    expect(detectDelimiter('名字\t备注\n张三\t已确认', '来源.txt')).toBe('\t');
    expect(parseDelimitedText('Name\tMemo\nAda\tpaid', '支出.tsv').rows).toEqual([
      ['Ada', 'paid'],
    ]);
  });

  it('读取转义引号、字段分隔符和带换行的引用字段', () => {
    expect(parseDelimitedText('Name,Note\r\nAda,"paid, pending"\r\nLin,"said ""yes""\nnext"', 'x.csv'))
      .toEqual({
        sourceLabel: 'x.csv',
        headers: ['Name', 'Note'],
        rows: [
          ['Ada', 'paid, pending'],
          ['Lin', 'said "yes"\nnext'],
        ],
      });
  });

  it('兼容 UTF-8 BOM、UTF-16 和无 BOM 的 GB18030 文本', () => {
    const encoded = new TextEncoder().encode('\uFEFF姓名,金额\n张三,12');
    expect(parseDelimitedFile(encoded.buffer, '中文.csv').rows[0]).toEqual(['张三', '12']);

    const utf16 = new Uint8Array([0xff, 0xfe, 0x4d, 0x00, 0x2c, 0x00, 0x31, 0x00]);
    expect(parseDelimitedFile(utf16.buffer, 'wide.csv').headers).toEqual(['M', '1']);

    const gb = Uint8Array.from([0xd5, 0xc5, 0xc8, 0xfd, 0x2c, 0x31]);
    expect(parseDelimitedFile(gb.buffer, 'legacy.csv').headers[0]).toBe('张三');
  });

  it('拒绝未闭合、错位引号和多余列，并让批次合并器接收的表结构无歧义', () => {
    for (const text of ['A,B\n"bad,1', 'A,B\na"b,c', 'A,B\n"a"x,b'])
      expect(() => parseDelimitedText(text, 'bad.csv')).toThrow();
    expect(() => parseDelimitedText('A,B\n1,2,3', 'bad.csv')).toThrow('多出单元格');
    expect(() => parseDelimitedText('', 'empty.csv')).toThrow('没有内容');
    expect(() => parseDelimitedText('A,A\n1,2', 'duplicate.csv')).toThrow('duplicate header');
  });

  it('导出逗号、引号和换行时遵循 CSV 转义，并标记公式风险而不篡改值', () => {
    expect(
      serializeSheetBatchCsv(['姓名', '备注'], [
        ['张三', '含,逗号'],
        ['李四', '说"好"\n第二行'],
      ]),
    ).toBe('\uFEFF姓名,备注\r\n张三,"含,逗号"\r\n李四,"说""好""\n第二行"\r\n');
    expect(
      formulaLikeTextCells({
        sourceLabel: 'safe.csv',
        headers: ['Name', '=HYPERLINK("x")'],
        rows: [['-2', 2], ['plain', '+cmd']],
      }),
    ).toBe(2);
  });
});
