import { createAdapter } from '../adapters';
import { createRuntime } from '../app/runtime';
import { userMessage } from '../core/errors';

// 仅本机开发验证入口。只允许本工程生成的、含固定标记的独立测试文件。
const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';
const hostName = params.get('host') ?? '';
const status = document.getElementById('status')!;
async function report(stage: string, details: Record<string, unknown> = {}) {
  status.textContent = `${stage}\n${JSON.stringify(details, null, 2)}`;
  const response = await fetch('/__native-report', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'X-Helper-Test': token },
    body: JSON.stringify({ host: hostName, stage, ...details }),
  });
  if (!response.ok) throw new Error('本地测试入口未启用，停止执行。');
}

async function run() {
  if (!import.meta.env.DEV || !token) throw new Error('仅用于本机开发验证');
  await report('page-loaded', { userAgent: navigator.userAgent });
  const adapter = await createAdapter(location.search);
  await report('host-connected', {
    kind: adapter.info.kind,
    capabilities: adapter.info.capabilities,
  });
  try {
    let selection = await adapter.readSelection();
    const expectedName =
      adapter.info.kind === 'text'
        ? 'OfficeHelper-Native-Test.docx'
        : 'OfficeHelper-Native-Test.xlsx';
    for (let attempt = 0; attempt < 50; attempt++) {
      const validName = selection.documentName === expectedName;
      const marker =
        adapter.info.kind === 'text'
          ? selection.text?.trim() === 'OFFICE_HELPER_NATIVE_TEST'
          : JSON.stringify(selection.values) === '[[120],[85],[45.5],[200]]';
      if (validName && marker) break;
      if (attempt === 49)
        throw new Error(
          `测试文件或选区标记不匹配，未执行修改。nameMatch=${validName}, markerMatch=${marker}, selectedLength=${selection.text?.length ?? 0}`,
        );
      await new Promise((resolve) => setTimeout(resolve, 600));
      selection = await adapter.readSelection();
    }
    await report('selection-verified', {
      kind: selection.kind,
      fontSize: selection.fontSize,
      rows: selection.values?.length,
    });
    const { runner } = createRuntime(adapter);
    if (selection.kind === 'text') {
      const originalSize = selection.fontSize;
      const task = await runner.prepare('text.enlarge');
      await runner.execute(task.id);
      const after = await adapter.readSelection();
      await report('resize-verified', { before: originalSize, after: after.fontSize });
      if (!after.fontSize || after.fontSize === originalSize) throw new Error('字号修改没有生效');
      await runner.undo();
      const restored = await adapter.readSelection();
      if (restored.fontSize !== originalSize || restored.text !== selection.text)
        throw new Error('恢复校验失败');
      await report('passed', {
        kind: 'text',
        before: originalSize,
        after: after.fontSize,
        restored: restored.fontSize,
      });
    } else {
      const task = await runner.prepare('sheet.sum');
      const result = await runner.execute(task.id);
      if (!result.message.includes('450.5')) throw new Error('合计不正确');
      await report('passed', { kind: 'sheet', total: 450.5, rows: selection.values?.length });
    }
  } finally {
    adapter.dispose();
  }
}

void run().catch(async (error) => {
  try {
    await report('failed', {
      message: userMessage(error),
      detail: error instanceof Error ? error.message : String(error),
      code: error?.code ?? '',
    });
  } catch {
    status.textContent = '测试入口已停止。请打开正常的小帮手入口。';
  }
});
