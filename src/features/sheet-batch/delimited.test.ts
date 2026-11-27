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
    expect(parseDelimitedText('Name\tMemo\nAda\tpaid', '支出.tsv').rows).toEqual([['Ada', 'paid']]);
  });

  it('读取转义引号、字段分隔符和带换行的引用字段', () => {
    expect(
      parseDelimitedText('Name,Note\r\nAda,"paid, pending"\r\nLin,"said ""yes""\nnext"', 'x.csv'),
    ).toEqual({
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

    const utf16be = new Uint8Array([0xfe, 0xff, 0x00, 0x41, 0x00, 0x2c, 0x00, 0x31]);
    expect(parseDelimitedFile(utf16be.buffer, 'wide-be.csv')).toMatchObject({
      headers: ['A', '1'],
      rows: [],
    });

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

  it('按记录编号处理带换行字段，并保留包含空值及尾部空字段的原始记录', () => {
    const source = parseDelimitedText('A,B\r\n"first\ncontinued",\r\n,tail\r\nlast,', 'rows.csv');
    expect(source.rows).toEqual([
      ['first\ncontinued', ''],
      ['', 'tail'],
      ['last', ''],
    ]);
  });

  it('解析时提前限制输入、数据记录、列数和含表头单元格总量', () => {
    expect(() => parseDelimitedText('A\n1\n2', 'rows.csv', { maximumRows: 1 })).toThrow(
      '记录数超过 1 条',
    );
    expect(() => parseDelimitedText('A,B\n1,2', 'columns.csv', { maximumColumns: 1 })).toThrow(
      '列数超过 1 列',
    );
    expect(() => parseDelimitedText('A,B\n1,2', 'cells.csv', { maximumCells: 3 })).toThrow(
      '单元格总数超过 3 格',
    );
    expect(() => parseDelimitedText('A\n12345', 'chars.csv', { maximumCharacters: 5 })).toThrow(
      '文字长度超过 5 个字符',
    );
    expect(() => parseDelimitedFile(new ArrayBuffer(4), 'bytes.csv', { maximumBytes: 3 })).toThrow(
      '超过 3 字节',
    );
    expect(() => parseDelimitedText('A\n1', 'invalid-limits.csv', { maximumRows: -1 })).toThrow(
      '必须是至少 0 的安全整数',
    );
  });

  it('拒绝 malformed UTF-16 字节且不以替换字符掩盖损坏', () => {
    const malformed = new Uint8Array([0xff, 0xfe, 0x41]);
    expect(() => parseDelimitedFile(malformed.buffer, 'malformed.csv')).toThrow();

    const undecodable = new Uint8Array([0x81, 0x30]);
    expect(() => parseDelimitedFile(undecodable.buffer, 'undecodable.csv')).toThrow('编码无法识别');
  });

  it('导出分隔符、引号和换行时遵循 CSV 转义，保留制表符并标记公式风险', () => {
    expect(
      serializeSheetBatchCsv(
        ['姓名', '备注', '制表符'],
        [
          ['张三', '含,逗号', '甲\t乙'],
          ['李四', '说"好"\n第二行', 'end'],
        ],
      ),
    ).toBe('\uFEFF姓名,备注,制表符\r\n张三,"含,逗号","甲\t乙"\r\n李四,"说""好""\n第二行",end\r\n');
    expect(
      formulaLikeTextCells({
        sourceLabel: 'safe.csv',
        headers: ['Name', '=HYPERLINK("x")'],
        rows: [
          ['-2', 2],
          ['-1.25', '+cmd'],
          ['-1e-3', '-1E+4'],
          ['-1-2', '-1.5+2'],
          ['-A1', '-SUM(A1)'],
          ['- note', 'plain'],
        ],
      }),
    ).toBe(7);
  });
});
