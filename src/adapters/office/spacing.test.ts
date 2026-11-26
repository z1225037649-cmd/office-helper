import { describe, expect, it } from 'vitest';
import { explicitMultipleSpacing } from './spacing';

const wrap = (content: string) =>
  `<pkg:package><pkg:part pkg:name="/word/document.xml"><w:document><w:body>${content}</w:body></w:document></pkg:part></pkg:package>`;
const p = (attrs: string) =>
  `<w:p><w:pPr><w:spacing ${attrs}/></w:pPr><w:r><w:t>通知</w:t></w:r></w:p>`;

describe('倍数行距的保守识别', () => {
  it('仅认可与原生数值一致的显式 auto 规则', () => {
    expect(
      explicitMultipleSpacing(
        wrap(p('w:line="240" w:lineRule="auto"') + p('w:lineRule="auto" w:line="360"')),
        [12, 18],
      ),
    ).toEqual([true, true]);
    expect(explicitMultipleSpacing(wrap(p('w:line="480" w:lineRule="auto"')), [12])).toEqual([
      false,
    ]);
  });
  it.each([
    'w:line="240" w:lineRule="exact"',
    'w:line="240" w:lineRule="atLeast"',
    'w:line="240"',
    'w:lineRule="auto"',
  ])('不猜测未知或固定行距：%s', (attrs) => {
    expect(explicitMultipleSpacing(wrap(p(attrs)), [12])).toEqual([false]);
  });
  it('样式继承、表格、修订及不匹配段落数量都拒绝', () => {
    const paragraph = p('w:line="240" w:lineRule="auto"');
    for (const xml of [
      wrap('<w:p><w:r><w:t>无直接格式</w:t></w:r></w:p>'),
      wrap(`<w:tbl>${paragraph}</w:tbl>`),
      wrap(`<w:ins>${paragraph}</w:ins>`),
      wrap(paragraph + paragraph),
      wrap(paragraph).replaceAll('w:', 'word:'),
    ])
      expect(explicitMultipleSpacing(xml, [12])).toEqual([false]);
  });
});
