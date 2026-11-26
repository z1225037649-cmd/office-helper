import { describe, expect, it } from 'vitest';
import { WpsAdapter } from '../../src/adapters/wps/adapter';

describe('WPS 读取适配契约', () => {
  it('支持单个数字单元格', async () => {
    const host = new WpsAdapter('sheet', () => ({
      ActiveWorkbook: { Name: '费用.xlsx' },
      ActiveSheet: { Name: '费用' },
      Selection: { Rows: { Count: 1 }, Columns: { Count: 1 }, Address: '$B$2', Value2: 120 },
    }));
    expect(await host.readSelection()).toMatchObject({ values: [[120]], targetId: '费用!$B$2' });
  });
  it('读取前限制范围大小，不加载整列', async () => {
    const host = new WpsAdapter('sheet', () => ({
      ActiveWorkbook: { Name: '费用.xlsx' },
      Selection: {
        Rows: { Count: 1048576 },
        Columns: { Count: 1 },
        get Value2(): never {
          throw new Error('不应读取');
        },
      },
    }));
    await expect(host.readSelection()).rejects.toMatchObject({ code: 'SELECTION_TOO_LARGE' });
  });
  it('缺失或不完整的数据不会伪装为空表格', async () => {
    const host = new WpsAdapter('sheet', () => ({
      ActiveWorkbook: { Name: '费用.xlsx' },
      Selection: { Rows: { Count: 2 }, Columns: { Count: 1 }, Address: 'A1:A2', Value2: [[12]] },
    }));
    await expect(host.readSelection()).rejects.toMatchObject({ code: 'UNAVAILABLE' });
  });
  it('没有稳定范围句柄时明确拒绝写入', async () => {
    const host = new WpsAdapter('text', () => ({
      ActiveDocument: { Name: '通知.docx' },
      Selection: { Text: '通知', Font: { Size: 18 }, Start: 0, End: 2 },
    }));
    expect(host.info.capabilities).toContain('text.resize');
    await expect(
      host.apply({ type: 'text.resize', size: 22 }, await host.readSelection()),
    ).rejects.toMatchObject({ code: 'UNAVAILABLE' });
  });
});
