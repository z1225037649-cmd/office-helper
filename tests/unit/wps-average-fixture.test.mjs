import { describe, expect, it } from 'vitest';
import { populateAverageWorksheet } from '../../scripts/prepare-wps-average-fixture.mjs';

describe('WPS 平均值专用表格样本生成', () => {
  it.each([
    '<worksheet><dimension ref="A1"/><sheetData/></worksheet>',
    '<worksheet><sheetData/></worksheet>',
    '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:dimension ref="A1"/><x:sheetData spans="1:2"/></x:worksheet>',
    '<worksheet><dimension ref="A1"/><sheetData><row r="1"/></sheetData></worksheet>',
  ])('生成标记和四个数值：%s', (xml) => {
    const result = populateAverageWorksheet(xml);
    expect(result).toContain('OFFICE_HELPER_NATIVE_TEST');
    for (const value of [120, 85, 45.5, 200]) {
      const escaped = String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      expect(result).toMatch(new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?v)>${escaped}</\\1>`));
    }
    expect(result).toContain('ref="A1:B5"');
    expect(result.match(/r="B[2-5]"/g)).toHaveLength(4);
  });

  it('拒绝缺少或多重的工作表数据区域', () => {
    expect(() => populateAverageWorksheet('<worksheet/>')).toThrow('exactly one');
    expect(() =>
      populateAverageWorksheet(
        '<worksheet><dimension ref="A1"/><sheetData/><sheetData/></worksheet>',
      ),
    ).toThrow('exactly one');
  });

  it('替换单引号形式的范围属性且不重复属性', () => {
    const result = populateAverageWorksheet(
      "<worksheet><dimension ref='A1'/><sheetData/></worksheet>",
    );
    expect(result.match(/\bref=/g)).toHaveLength(1);
    expect(result).toContain('ref="A1:B5"');
  });

  it('在 sheetPr 后插入缺失的范围节点', () => {
    const result = populateAverageWorksheet('<worksheet><sheetPr/><sheetData/></worksheet>');
    expect(result.indexOf('<sheetPr')).toBeLessThan(result.indexOf('<dimension'));
    expect(result.indexOf('<dimension')).toBeLessThan(result.indexOf('<sheetData'));
  });

  it('忽略注释中的伪 sheetPr 标签并保留原注释', () => {
    const result = populateAverageWorksheet('<worksheet><!--<sheetPr/>--><sheetData/></worksheet>');
    expect(result).toContain('<!--<sheetPr/>-->');
    expect(result.match(/<dimension\b/g)).toHaveLength(1);
    expect(result.indexOf('<dimension')).toBeLessThan(result.indexOf('<sheetData'));
    expect(result).not.toContain('<sheetPr/><dimension');
  });

  it('不把注释中的伪 sheetData 标签算作工作表数据节点', () => {
    const result = populateAverageWorksheet(
      '<worksheet><!--<sheetData/>--><sheetData/></worksheet>',
    );
    expect(result).toContain('<!--<sheetData/>-->');
    expect(result.match(/<sheetData\b/g)).toHaveLength(2);
    expect(result).toContain('<v>120</v>');
  });

  it('保留 SpreadsheetML 命名空间并使用声明的前缀', () => {
    const result = populateAverageWorksheet(
      '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/></x:worksheet>',
    );
    expect(result).toContain('xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    expect(result).toContain('<x:dimension ref="A1:B5"/>');
    expect(result).toContain('<x:v>120</x:v>');
  });
});
