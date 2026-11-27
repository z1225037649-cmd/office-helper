import { describe, expect, it } from 'vitest';
import { features } from '../../src/features';
import { LocalIntentProvider } from '../../src/services/intent';

const intent = new LocalIntentProvider();

describe('扩展功能声明的短句', () => {
  const extendedIds = new Set([
    'sheet.autofit',
    'sheet.wrap',
    'sheet.header',
    'sheet.print',
    'document.saveCopy',
    'document.exportPdf',
    'sheet.checkRecords',
    'document.notice',
  ]);
  const phrases = features
    .filter((feature) => extendedIds.has(feature.id))
    .flatMap((feature) => feature.phrases.map((input) => ({ input, featureId: feature.id })));

  it.each(phrases)('$featureId 保留短句及礼貌前缀：$input', async ({ input, featureId }) => {
    for (const request of [input, `请帮我${input}一下，谢谢。`, `麻烦你${input}即可。`]) {
      expect(await intent.resolve(request, features), request).toEqual({
        type: 'match',
        featureId,
      });
    }
  });
});

describe('表格格式指令保留完整意图', () => {
  it.each([
    ['自动调整列宽', 'sheet.autofit'],
    ['请帮我自动调整列宽一下', 'sheet.autofit'],
    ['把选中的列自动列宽', 'sheet.autofit'],
    ['麻烦你将当前选区列宽自适应即可。', 'sheet.autofit'],
    ['自动换行', 'sheet.wrap'],
    ['请把选中的单元格自动换行', 'sheet.wrap'],
    ['单元格换行', 'sheet.wrap'],
    ['麻烦你将当前选区单元格换行即可。', 'sheet.wrap'],
    ['启用换行', 'sheet.wrap'],
    ['突出表头', 'sheet.header'],
    ['请把当前选区加粗表头', 'sheet.header'],
    ['设置表头', 'sheet.header'],
  ])('支持完整当前选区请求：%s', async (input, featureId) => {
    expect(await intent.resolve(input!, features)).toEqual({ type: 'match', featureId });
  });

  it.each([
    '关闭自动换行',
    '禁用自动换行',
    '不自动换行',
    '请不要把当前选区自动换行',
    '请把当前选区的自动换行关闭，谢谢。',
    '不启用换行',
    '取消单元格换行',
    '只把第二行自动换行',
    '第2行自动换行',
    '把选中单元格的第二行自动换行',
    '请将A1:B2自动换行',
    '自动换行只对第二行',
    '仅对可见单元格自动换行',
    '自动换行但排除隐藏行',
    '自动调整列宽到20磅',
    '把列宽自适应设为20',
    '调整列宽为三十',
    '请把当前选区自动调整列宽到２０磅，谢谢。',
    '请把选中的列自动列宽20pt',
    '请把当前选区列宽自适应设为三十',
    '请将第二列自动调整列宽',
    '只调整列宽不影响选区外的单元格',
    '去掉加粗表头',
    '移除加粗表头',
    '请取消加粗表头',
    '不加粗表头',
    '请把当前选区加粗表头去掉，谢谢。',
    '请把选中单元格的第二行加粗表头',
    '只对第二行突出表头',
    '突出表头字号20磅',
    '突出表头并变红',
    '突出表头是什么意思',
    '请把当前选区加粗表头成红色',
    '我昨天听说了加粗表头',
  ])('不支持的状态、范围、数值或附加要求先澄清：%s', async (input) => {
    expect(await intent.resolve(input, features)).toMatchObject({ type: 'clarify' });
  });

  it.each(['加粗表头并加粗', '加粗并加粗表头', '加粗表头后取消加粗', '突出表头但不加粗'])(
    '表头之外的加粗要求仍需逐项说明：%s',
    async (input) => {
      expect(await intent.resolve(input, features)).toEqual({
        type: 'choices',
        featureIds: ['text.bold', 'sheet.header'],
      });
    },
  );

  it('表头短句重叠消除后仍保留其他独立功能', async () => {
    expect(await intent.resolve('加粗表头并居中', features)).toEqual({
      type: 'choices',
      featureIds: ['paragraph.center', 'sheet.header'],
    });
    expect(await intent.resolve('加粗表头并自动换行', features)).toEqual({
      type: 'choices',
      featureIds: ['sheet.wrap', 'sheet.header'],
    });
  });
});

describe('明确的文字格式指令', () => {
  it.each(['请加粗', '把选中文字设置为粗体'])('加粗：%s', async (input) => {
    expect(await intent.resolve(input, features)).toEqual({
      type: 'match',
      featureId: 'text.bold',
      parameters: { bold: true },
    });
  });

  it.each(['取消加粗', '请取消粗体，谢谢。'])('取消加粗：%s', async (input) => {
    expect(await intent.resolve(input, features)).toEqual({
      type: 'match',
      featureId: 'text.bold',
      parameters: { bold: false },
    });
  });

  it.each(['段落居中', '请把选中的文字居中对齐。'])('居中：%s', async (input) => {
    expect(await intent.resolve(input, features)).toEqual({
      type: 'match',
      featureId: 'paragraph.center',
    });
  });

  it.each([
    ['行距设为1.5倍', 1.5],
    ['请把行距调整到 2 倍。', 2],
    ['设置为1.25倍行距', 1.25],
    ['1倍行距', 1],
    ['行距3倍', 3],
    ['行距设为１.５倍', 1.5],
  ])('行距保留小数：%s', async (input, lineSpacingMultiple) => {
    expect(await intent.resolve(String(input), features)).toEqual({
      type: 'match',
      featureId: 'paragraph.spacing',
      parameters: { lineSpacingMultiple },
    });
  });

  it.each([
    ['首行缩进24磅', 24],
    ['首行缩进设为0pt', 0],
    ['首行缩进 0.74 厘米', 21],
    ['请把首行缩进改为1cm。', 28.35],
    ['首行缩进5.08厘米', 144],
    ['首行缩进24.024磅', 24],
  ])('首行缩进换算为磅：%s', async (input, firstLineIndentPoints) => {
    expect(await intent.resolve(String(input), features)).toEqual({
      type: 'match',
      featureId: 'paragraph.indent',
      parameters: { firstLineIndentPoints },
    });
  });

  it('字号小数不会被删除小数点，未指定字号时仍支持原来的放大', async () => {
    expect(await intent.resolve('把字放大到30.5磅', features)).toEqual({
      type: 'match',
      featureId: 'text.enlarge',
      parameters: { fontSize: 30.5 },
    });
    expect(await intent.resolve('请帮我把字放大一点。', features)).toEqual({
      type: 'match',
      featureId: 'text.enlarge',
    });
  });
});

describe('不完整或歧义的格式指令先澄清', () => {
  it.each([
    '不要加粗',
    '不加粗',
    '请别把文字加粗',
    '不需要粗体',
    '不要取消加粗',
    '别取消加粗',
    '请勿取消粗体',
    '无须粗体',
    '不能加粗',
    '避免加粗',
    '不要居中',
    '不居中',
    '取消居中',
    '取消行距',
    '不要把字放大',
    '取消首行缩进',
  ])('否定或不支持的取消不会执行相反操作：%s', async (input) => {
    expect(await intent.resolve(input, features)).toMatchObject({ type: 'clarify' });
  });

  it.each([
    '行距',
    '行距大一点',
    '行距设为',
    '行距20磅',
    '行距固定20磅',
    '固定行距为1.5倍',
    '行距1.5',
    '行距0.5倍',
    '行距3.1倍',
    '行距1.234倍',
    '把字放大到30.3磅',
    '行距-1.5倍',
    '行距−1.5倍',
    '行距1.5倍厘米',
    '行距1.5或2倍',
    '行距1.5倍2倍',
    '行距1，5倍',
    '行距1 5倍',
    '行距1..5倍',
    '行距一倍',
    '首行缩进',
    '缩进一点',
    '缩进24磅',
    '首行缩进2字符',
    '首行缩进二字符',
    '首行缩进24',
    '首行缩进-24磅',
    '首行缩进145磅',
    '首行缩进5.09厘米',
    '首行缩进1英寸',
    '首行缩进1cm磅',
    '首行缩进1cm2磅',
    '首行缩进24或者30磅',
    '悬挂首行缩进24磅',
    '加粗2倍',
    '加粗三倍',
    '加粗再取消加粗',
    '加粗并变成红色',
    '加粗且删除文字',
    '加粗成斜体',
    '加粗以后保存',
    '粗体是什么意思',
    '把所有文字加粗',
    '居中2厘米',
    '居中然后斜体',
    '居中后保存',
    '居中成斜体',
    '居中删除文字',
    '居中是什么意思',
    '全部段落居中',
    '把字放大成红色',
    '先删除文字再把字放大',
    '删除文字行距设为1.5倍',
    '删除文字首行缩进24磅',
  ])('未知数值、单位或多步要求不会使用默认值：%s', async (input) => {
    expect(await intent.resolve(input, features)).toMatchObject({ type: 'clarify' });
  });

  it.each([
    '把字放大到-30磅',
    '把字放大到−30磅',
    '把字放大到+30磅',
    '把字放大到30磅厘米',
    '把字放大到30cmpt',
    '把字放大到30像素',
    '把字放大到30pt50pt',
    '把字放大到30或50磅',
    '把字放大到3，0磅',
    '把字放大到3 0磅',
    '把字放大到3e1磅',
    '把字放大到30..5磅',
    '把字放大到三十磅',
    '把字放大一号',
  ])('字号不会截取歧义片段作为合法数字：%s', async (input) => {
    expect(await intent.resolve(input, features)).toMatchObject({ type: 'clarify' });
  });

  it('多功能指令不返回丢失参数的单个匹配，需逐项重新说明', async () => {
    expect(await intent.resolve('把字放大到30磅，然后行距设为1.5倍', features)).toEqual({
      type: 'choices',
      featureIds: ['text.enlarge', 'paragraph.spacing'],
    });
    expect(await intent.resolve('加粗并居中', features)).toMatchObject({ type: 'choices' });
  });

  it('未开放的功能仍不可匹配，并尊重中止信号', async () => {
    expect(await intent.resolve('加粗', [])).toEqual({ type: 'unknown' });
    const controller = new AbortController();
    controller.abort();
    await expect(intent.resolve('加粗', features, controller.signal)).rejects.toThrow();
  });
});
