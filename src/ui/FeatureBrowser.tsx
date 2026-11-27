import { useId, useRef, useState } from 'react';
import {
  AlignCenter,
  Bold,
  Calculator,
  ChevronRight,
  IndentIncrease,
  ListCollapse,
  ScanText,
  Search,
  Sigma,
  Columns3,
  WrapText,
  Printer,
  Copy,
  FileDown,
  FilePlus2,
  ScanSearch,
  Type,
  X,
  type LucideIcon,
} from 'lucide-react';
import type { FeatureDefinition } from '../core/contracts';

const featureLabels: Record<string, { label: string; Icon: LucideIcon }> = {
  'text.enlarge': { label: '调整字号', Icon: Type },
  'text.documentStyles': { label: '统一全文标题', Icon: ListCollapse },
  'document.bidChecklist': { label: '检查标书要求', Icon: ScanSearch },
  'text.bold': { label: '加粗文字', Icon: Bold },
  'paragraph.center': { label: '段落居中', Icon: AlignCenter },
  'paragraph.spacing': { label: '设置行距', Icon: ListCollapse },
  'paragraph.indent': { label: '首行缩进', Icon: IndentIncrease },
  'sheet.sum': { label: '计算合计', Icon: Sigma },
  'sheet.average': { label: '计算平均值', Icon: Calculator },
  'selection.inspect': { label: '查看选区', Icon: ScanText },
  'sheet.autofit': { label: '自动调整列宽', Icon: Columns3 },
  'sheet.wrap': { label: '单元格自动换行', Icon: WrapText },
  'sheet.header': { label: '强调选区表头', Icon: Bold },
  'sheet.print': { label: '打印前检查', Icon: Printer },
  'sheet.checkRecords': { label: '找重复和漏填', Icon: ScanSearch },
  'document.saveCopy': { label: '另存一份', Icon: Copy },
  'document.exportPdf': { label: '保存成 PDF', Icon: FileDown },
  'document.notice': { label: '生成通知', Icon: FilePlus2 },
};

const fallbackIcons = { text: Type, sum: Sigma, inspect: ScanText };
const groups = [
  { id: 'text', label: '文字' },
  { id: 'paragraph', label: '段落' },
  { id: 'sheet', label: '表格' },
  { id: 'document', label: '文件' },
  { id: 'other', label: '其他' },
] as const;

function featureGroup(id: string) {
  if (id.startsWith('text.')) return 'text';
  if (id.startsWith('paragraph.')) return 'paragraph';
  if (id.startsWith('sheet.')) return 'sheet';
  if (id.startsWith('document.')) return 'document';
  return 'other';
}

export function FeatureBrowser({
  features,
  busy,
  unavailableReason,
  onPrepare,
  onOpenBatchTool,
}: {
  features: readonly FeatureDefinition[];
  busy: boolean;
  unavailableReason: (feature: FeatureDefinition) => string | undefined;
  onPrepare: (id: string) => void;
  onOpenBatchTool: () => void;
}) {
  const [search, setSearch] = useState('');
  const searchRef = useRef<HTMLInputElement>(null);
  const headingId = useId();
  const entries = features.map((feature) => ({ feature, reason: unavailableReason(feature) }));
  // The caller supplies features already filtered to the current document kind.
  const isSheet = features.some(
    (feature) => feature.kinds.includes('sheet') && !feature.kinds.includes('text'),
  );
  const preferredIds = isSheet
    ? ['sheet.sum', 'sheet.average', 'selection.inspect']
    : ['text.enlarge', 'text.bold', 'paragraph.center', 'selection.inspect'];
  const quickFeatures = [
    ...preferredIds.flatMap((id) => {
      const entry = entries.find(({ feature }) => feature.id === id);
      return entry ? [entry] : [];
    }),
    ...entries.filter(({ feature }) => !preferredIds.includes(feature.id)),
  ].slice(0, 4);
  const query = search.trim().toLocaleLowerCase();
  const matchingFeatures = entries.filter(({ feature }) =>
    [
      feature.title,
      featureLabels[feature.id]?.label ?? feature.title,
      feature.description,
      ...feature.phrases,
    ].some((text) => text.toLocaleLowerCase().includes(query)),
  );

  function renderFeature({ feature, reason }: (typeof entries)[number], catalog = false) {
    const { label, Icon } = featureLabels[feature.id] ?? {
      label: feature.title,
      Icon: fallbackIcons[feature.icon],
    };
    const description = reason ?? feature.description;

    return (
      <button
        key={feature.id}
        type="button"
        className={catalog ? 'catalog-item' : 'quick-feature'}
        disabled={busy || !!reason}
        aria-label={label}
        aria-description={description}
        title={description}
        onClick={() => onPrepare(feature.id)}
      >
        <Icon size={16} aria-hidden="true" />
        {catalog ? (
          <span className="catalog-item-copy">
            <span className="feature-label">{label}</span>
            <span className="catalog-item-description">{description}</span>
          </span>
        ) : (
          <span className="feature-label">{label}</span>
        )}
      </button>
    );
  }

  return (
    <>
      <section className="feature-section" aria-labelledby={headingId}>
        <div className="feature-section-heading">
          <h2 id={headingId}>常用功能</h2>
          <span className="feature-count">{quickFeatures.length}</span>
        </div>
        <div className="quick-features">{quickFeatures.map((entry) => renderFeature(entry))}</div>
        {!features.length && <p className="feature-empty">当前没有可用功能。</p>}
      </section>
      <details className="feature-catalog">
        <summary>
          <ChevronRight size={16} className="catalog-chevron" aria-hidden="true" />
          <span>全部功能</span>
          <span className="feature-count">{features.length}</span>
        </summary>
        <div className="feature-catalog-content">
          <div className="feature-search">
            <Search size={16} aria-hidden="true" />
            <input
              ref={searchRef}
              type="search"
              value={search}
              placeholder="搜索功能"
              aria-label="搜索功能"
              autoComplete="off"
              onChange={(event) => setSearch(event.target.value)}
            />
            {search && (
              <button
                type="button"
                className="feature-search-clear"
                aria-label="清空搜索"
                title="清空搜索"
                onClick={() => {
                  setSearch('');
                  searchRef.current?.focus();
                }}
              >
                <X size={16} aria-hidden="true" />
              </button>
            )}
          </div>
          {groups.map((group) => {
            const items = matchingFeatures.filter(
              ({ feature }) => featureGroup(feature.id) === group.id,
            );
            if (!items.length) return null;

            return (
              <section
                key={group.id}
                className="catalog-group"
                aria-labelledby={`${headingId}-${group.id}`}
              >
                <h3 id={`${headingId}-${group.id}`} className="catalog-group-heading">
                  {group.label}
                </h3>
                <div className="catalog-items">
                  {items.map((entry) => renderFeature(entry, true))}
                </div>
              </section>
            );
          })}
          {isSheet && (
            <section className="catalog-group" aria-labelledby={`${headingId}-batch`}>
              <h3 id={`${headingId}-batch`} className="catalog-group-heading">批量文件</h3>
              <div className="catalog-items">
                <button
                  type="button"
                  className="catalog-item"
                  disabled={busy}
                  onClick={onOpenBatchTool}
                >
                  <ScanSearch size={16} aria-hidden="true" />
                  <span className="catalog-item-copy">
                    <span className="feature-label">批量合并与数据检查</span>
                    <span className="catalog-item-description">导入多个 CSV／TSV，按列名合并、查看异常候选并下载结果</span>
                  </span>
                </button>
              </div>
            </section>
          )}
          {!matchingFeatures.length && (
            <p className="catalog-empty" role="status">
              {features.length
                ? '没有匹配的功能，试试其他关键词，或清空搜索。'
                : '当前没有可用功能。'}
            </p>
          )}
        </div>
      </details>
    </>
  );
}
