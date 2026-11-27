import { describe, expect, it } from 'vitest';
import { MockAdapter } from '../../src/adapters/mock/adapter';
import { FeatureRegistry } from '../../src/core/registry';
import { TaskRunner } from '../../src/core/runner';
import { features } from '../../src/features';

describe('全文标题样式工作流', () => {
  it('先展示范围和识别结果，确认修改后可恢复', async () => {
    const host = new MockAdapter('text');
    const runner = new TaskRunner(host, new FeatureRegistry(features));
    const before = await host.readSelection({ includeDocumentParagraphs: true });
    const preview = await runner.prepare('text.documentStyles');

    expect(preview.target).toContain('全文');
    expect(preview.summary).toContain('一级 2 个');
    expect(await host.readSelection({ includeDocumentParagraphs: true })).toEqual(before);

    const result = await runner.execute(preview.id);
    expect(result.canUndo).toBe(true);
    expect(
      (await host.readSelection({ includeDocumentParagraphs: true })).documentParagraphs?.map(
        (paragraph) => paragraph.style,
      ),
    ).toEqual(['Heading1', 'Heading1', 'Heading2', 'Normal', 'Heading1', 'Normal', 'Normal']);

    await runner.undo();
    expect(await host.readSelection({ includeDocumentParagraphs: true })).toEqual(before);
  });
});
