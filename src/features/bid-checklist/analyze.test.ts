import { describe, expect, it } from 'vitest';
import { analyzeBidChecklist } from './analyze';

describe('analyzeBidChecklist', () => {
  it('extracts likely requirement lines with exact excerpts, offsets, categories, and empty status', () => {
    const source = [
      'Invitation for tenders',
      '1. The bidder must provide evidence of valid insurance.',
      '2. The solution shall meet the security requirements.',
      '3. Prices shall be quoted in Hong Kong dollars.',
      '4. Tender submissions must be received no later than 5pm on 1 December.',
      '5. Submit signed form A using the portal.',
      'Contact the buyer for further information.',
    ].join('\r\n');

    const candidates = analyzeBidChecklist(source);

    expect(
      candidates.map(({ originalExcerpt, category, status, isHeuristicCandidate }) => ({
        originalExcerpt,
        category,
        status,
        isHeuristicCandidate,
      })),
    ).toEqual([
      {
        originalExcerpt: '1. The bidder must provide evidence of valid insurance.',
        category: 'qualification',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
      {
        originalExcerpt: '2. The solution shall meet the security requirements.',
        category: 'technical',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
      {
        originalExcerpt: '3. Prices shall be quoted in Hong Kong dollars.',
        category: 'commercial',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
      {
        originalExcerpt: '4. Tender submissions must be received no later than 5pm on 1 December.',
        category: 'deadline',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
      {
        originalExcerpt: '5. Submit signed form A using the portal.',
        category: 'submission',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
    ]);

    for (const candidate of candidates) {
      expect(source.slice(candidate.startOffset, candidate.endOffset)).toBe(
        candidate.originalExcerpt,
      );
    }
  });

  it('识别常见中文标书要求并分类，保留原文偏移', () => {
    const source = [
      '第一章 采购要求',
      '投标人必须提供有效营业执照和类似项目业绩。',
      '所投设备应满足技术参数及安全标准。',
      '本项目报价须包含税费及运输费用。',
      '请于2026年12月1日17时前提交投标文件。',
      '提交加盖公章的资格审查表。',
      '如有问题可联系采购人。',
    ].join('\n');

    const candidates = analyzeBidChecklist(source);

    expect(candidates.map(({ originalExcerpt, category }) => [originalExcerpt, category])).toEqual([
      ['投标人必须提供有效营业执照和类似项目业绩。', 'qualification'],
      ['所投设备应满足技术参数及安全标准。', 'technical'],
      ['本项目报价须包含税费及运输费用。', 'commercial'],
      ['请于2026年12月1日17时前提交投标文件。', 'deadline'],
      ['提交加盖公章的资格审查表。', 'submission'],
    ]);
    for (const candidate of candidates)
      expect(source.slice(candidate.startOffset, candidate.endOffset)).toBe(
        candidate.originalExcerpt,
      );
  });

  it('uses stable, exclusive offsets with CRLF, Unicode, and leading or trailing whitespace', () => {
    const source = '🌏 Context\r\n  Provide a signed form.  \n   \r\n  Deadline: 5pm.\r';

    expect(analyzeBidChecklist(source)).toEqual([
      {
        originalExcerpt: '  Provide a signed form.  ',
        startOffset: source.indexOf('  Provide'),
        endOffset: source.indexOf('  Provide') + '  Provide a signed form.  '.length,
        category: 'submission',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
      {
        originalExcerpt: '  Deadline: 5pm.',
        startOffset: source.indexOf('  Deadline'),
        endOffset: source.length - 1,
        category: 'deadline',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
    ]);
  });

  it('recognizes numbered imperative instructions but ignores ordinary prose and headings', () => {
    expect(
      analyzeBidChecklist(
        [
          'Qualification Requirements',
          'The procurement meeting is scheduled for Monday.',
          '(a) Attach the completed declaration.',
          '* Please confirm the authorized signatory.',
          'Further details are available on request.',
        ].join('\n'),
      ).map(({ originalExcerpt, category }) => ({ originalExcerpt, category })),
    ).toEqual([
      { originalExcerpt: '(a) Attach the completed declaration.', category: 'submission' },
      { originalExcerpt: '* Please confirm the authorized signatory.', category: 'other' },
    ]);
  });

  it('recognizes Chinese numbered instructions and keeps deadlines ahead of submission actions', () => {
    const candidates = analyzeBidChecklist(
      [
        '一、提交经签署的承诺书。',
        '（二）于2026年12月1日前提交投标文件。',
        '第3条 请在开标前提供样品。',
        '1) Submit the signed form by the deadline.',
      ].join('\r\n'),
    );

    expect(candidates.map(({ originalExcerpt, category }) => [originalExcerpt, category])).toEqual([
      ['一、提交经签署的承诺书。', 'submission'],
      ['（二）于2026年12月1日前提交投标文件。', 'deadline'],
      ['第3条 请在开标前提供样品。', 'deadline'],
      ['1) Submit the signed form by the deadline.', 'deadline'],
    ]);
  });

  it('ignores explicit negative requirements and ordinary explanatory wording', () => {
    const source = [
      '本项仅供参考，无需提交任何证明。',
      '供应商不必填写此栏目。',
      '投标人不需要提供纸本副本。',
      'No separate certificate is required.',
      'The bidder is not required to attach this optional schedule.',
      'This item is not mandatory.',
      'Please note that no additional form is required.',
      '投标人必须提供有效营业执照。',
      '投标人不得转让本合同。',
    ].join('\n');

    expect(analyzeBidChecklist(source).map(({ originalExcerpt }) => originalExcerpt)).toEqual([
      '投标人必须提供有效营业执照。',
      '投标人不得转让本合同。',
    ]);
  });

  it('returns no candidates for empty text, headings, or blank lines', () => {
    expect(analyzeBidChecklist('')).toEqual([]);
    expect(analyzeBidChecklist('\n\r\nTechnical Specifications\n\t\n')).toEqual([]);
  });

  it('keeps category fallback cautious for explicit requirements without a category signal', () => {
    expect(analyzeBidChecklist('The contractor shall comply with the stated conditions.')).toEqual([
      {
        originalExcerpt: 'The contractor shall comply with the stated conditions.',
        startOffset: 0,
        endOffset: 'The contractor shall comply with the stated conditions.'.length,
        category: 'other',
        status: 'unfilled',
        isHeuristicCandidate: true,
      },
    ]);
  });

  it('returns fresh deterministic results without modifying the source', () => {
    const source = 'Submit the bid by the stated deadline.';
    const first = analyzeBidChecklist(source);
    const second = analyzeBidChecklist(source);

    expect(first).toEqual(second);
    expect(first).not.toBe(second);
    expect(source).toBe('Submit the bid by the stated deadline.');
  });

  it('scans a long line without truncating its matching requirement or UTF-16 offsets', () => {
    const source = `${'背景'.repeat(40_000)} 投标人必须提供资质证明。`;
    const candidates = analyzeBidChecklist(source);
    expect(candidates).toHaveLength(1);
    const candidate = candidates[0]!;
    const expectedExcerpt = '投标人必须提供资质证明。';

    expect(candidate).toEqual({
      originalExcerpt: source,
      startOffset: 0,
      endOffset: source.length,
      category: 'qualification',
      status: 'unfilled',
      isHeuristicCandidate: true,
    });
    expect(source.slice(candidate.startOffset, candidate.endOffset)).toBe(source);
    expect(candidate.originalExcerpt.endsWith(expectedExcerpt)).toBe(true);
  });
});
