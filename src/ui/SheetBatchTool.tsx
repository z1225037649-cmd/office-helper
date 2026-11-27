import { useMemo, useRef, useState } from 'react';
import { Download, FileSpreadsheet, Trash2, X } from 'lucide-react';
import {
  analyzeSheetNormalization,
  createSheetBatchPlan,
  formulaLikeTextCells,
  parseDelimitedFile,
  serializeSheetBatchCsv,
  type NormalizationFinding,
  type SheetBatchSource,
} from '../features/sheet-batch';

interface BatchInput {
  id: string;
  source: SheetBatchSource;
  size: number;
}

const maximumFiles = 10;
const maximumFileBytes = 10 * 1024 * 1024;
const maximumTotalBytes = 50 * 1024 * 1024;
const maximumRows = 100_000;
const maximumCells = 1_000_000;
const maximumColumns = 200;

function uniqueLabel(fileName: string, inputs: readonly BatchInput[]): string {
  const labels = new Set(inputs.map(({ source }) => source.sourceLabel));
  if (!labels.has(fileName)) return fileName;
  const dot = fileName.lastIndexOf('.');
  const stem = dot > 0 ? fileName.slice(0, dot) : fileName;
  const extension = dot > 0 ? fileName.slice(dot) : '';
  for (let index = 2; ; index += 1) {
    const candidate = `${stem} (${index})${extension}`;
    if (!labels.has(candidate)) return candidate;
  }
}

function findingLocation(finding: NormalizationFinding): string {
  return finding.rowIndex === null
    ? `整列「${finding.columnName}」`
    : `第 ${finding.rowIndex + 1} 行「${finding.columnName}」`;
}

export function SheetBatchTool({
  busy,
  onClose,
}: {
  busy: boolean;
  onClose: () => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [inputs, setInputs] = useState<BatchInput[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [status, setStatus] = useState('');

  const plan = useMemo(
    () => (inputs.length ? createSheetBatchPlan(inputs.map(({ source }) => source)) : undefined),
    [inputs],
  );
  const analysis = useMemo(
    () => (inputs.length ? analyzeSheetNormalization(inputs.map(({ source }) => source)) : undefined),
    [inputs],
  );
  const combinedSource = useMemo(
    () =>
      plan
        ? {
            sourceLabel: 'merged.csv',
            headers: [...plan.columns, '来源文件', '原始行号'],
            rows: plan.rows.map(({ values, sourceLabel, originalRowIndex }) => [
              ...values,
              sourceLabel,
              originalRowIndex + 1,
            ]),
          }
        : undefined,
    [plan],
  );
  const unsafeCells = combinedSource ? formulaLikeTextCells(combinedSource) : 0;
  const previewRows = plan?.rows.slice(0, 8) ?? [];
  const previewFindings = analysis?.findings.slice(0, 30) ?? [];

  async function addFiles(fileList: FileList | null) {
    if (!fileList?.length) return;
    setLoading(true);
    setError('');
    setStatus('');
    try {
      const incoming = Array.from(fileList);
      if (inputs.length + incoming.length > maximumFiles)
        throw new TypeError(`一次最多处理 ${maximumFiles} 个文件。`);
      if (incoming.some((file) => !/\.(?:csv|tsv)$/iu.test(file.name)))
        throw new TypeError('请选择 CSV 或 TSV 文件。Excel 工作簿请先另存为 CSV 或制表符分隔文件。');
      if (incoming.some((file) => file.size > maximumFileBytes))
        throw new TypeError('单个文件不能超过 10 MB。');
      const nextBytes = inputs.reduce((sum, item) => sum + item.size, 0) +
        incoming.reduce((sum, file) => sum + file.size, 0);
      if (nextBytes > maximumTotalBytes) throw new TypeError('所有文件合计不能超过 50 MB。');

      const parsed: BatchInput[] = [];
      let labels = [...inputs];
      for (const file of incoming) {
        const label = uniqueLabel(file.name, labels);
        const source = parseDelimitedFile(await file.arrayBuffer(), label);
        parsed.push({ id: crypto.randomUUID(), source, size: file.size });
        labels = [...labels, parsed[parsed.length - 1]!];
      }

      const next = [...inputs, ...parsed];
      const rowCount = next.reduce((sum, item) => sum + item.source.rows.length, 0);
      const columnCount = new Set(next.flatMap((item) => item.source.headers)).size;
      if (rowCount > maximumRows)
        throw new TypeError(`数据行数超过 ${maximumRows.toLocaleString()} 行，请拆成较小批次。`);
      if (columnCount > maximumColumns)
        throw new TypeError(`合并后的列数超过 ${maximumColumns} 列，请先核对表头。`);
      if (rowCount * columnCount > maximumCells)
        throw new TypeError('合并后的数据量超过本工具的安全预览上限，请拆成较小批次。');

      setInputs(next);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : '文件读取失败，请另存为 UTF-8 CSV 后重试。');
    } finally {
      setLoading(false);
      if (inputRef.current) inputRef.current.value = '';
    }
  }

  function downloadMerged() {
    if (!plan || !combinedSource || unsafeCells > 0 || busy || loading) return;
    const csv = serializeSheetBatchCsv(
      combinedSource.headers,
      combinedSource.rows,
    );
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = '合并结果.csv';
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    setStatus('已交给浏览器下载，请在下载列表中核对文件名和内容。');
  }

  return (
    <section className="sheet-batch-tool" aria-label="表格批量合并与检查">
      <div className="sheet-batch-heading">
        <span className="assistant-symbol"><FileSpreadsheet size={18} /></span>
        <div>
          <h2>批量合并与数据检查</h2>
          <p>导入多个 CSV／TSV，按完全相同的列名对齐，先预览再下载。</p>
        </div>
        <button type="button" className="icon-button" aria-label="关闭批量工具" onClick={onClose}>
          <X size={18} />
        </button>
      </div>

      <p className="sheet-batch-note">
        支持 UTF-8、UTF-16 和 GB18030 编码。最多 10 个文件、每个 10 MB、合计 50 MB；结果只在浏览器生成，不会改动当前工作簿。含公式样式文字时会暂停 CSV 下载，避免打开文件时触发公式。
      </p>
      <input
        ref={inputRef}
        className="visually-hidden-input"
        type="file"
        accept=".csv,.tsv,text/csv,text/tab-separated-values"
        multiple
        aria-label="选择 CSV 或 TSV 文件"
        disabled={busy || loading}
        onChange={(event) => void addFiles(event.currentTarget.files)}
      />
      <div className="preview-buttons">
        <button
          type="button"
          className="primary-button"
          disabled={busy || loading || inputs.length >= maximumFiles}
          onClick={() => inputRef.current?.click()}
        >
          {loading ? '正在读取…' : '选择 CSV／TSV 文件'}
        </button>
        {plan && (
          <button
            type="button"
            className="secondary-button"
            disabled={busy || loading || unsafeCells > 0}
            title={unsafeCells ? '请先处理以公式符号开头的文字单元格，再导出 CSV。' : undefined}
            onClick={downloadMerged}
          >
            <Download size={16} />
            下载合并结果
          </button>
        )}
      </div>

      {error && <p className="sheet-batch-error" role="alert">{error}</p>}
      {status && <p className="sheet-batch-status" role="status">{status}</p>}
      {!!inputs.length && (
        <ul className="sheet-batch-files" aria-label="已导入文件">
          {inputs.map(({ id, source, size }) => (
            <li key={id}>
              <span>{source.sourceLabel} · {source.rows.length.toLocaleString()} 行 · {(size / 1024).toFixed(0)} KB</span>
              <button
                type="button"
                className="icon-button"
                aria-label={`移除 ${source.sourceLabel}`}
                disabled={busy || loading}
                onClick={() => {
                  setInputs((items) => items.filter((item) => item.id !== id));
                  setStatus('');
                }}
              >
                <Trash2 size={16} />
              </button>
            </li>
          ))}
        </ul>
      )}

      {plan && analysis && (
        <>
          <div className="sheet-batch-summary" role="status">
            <strong>{plan.rowCountSummary.sourceCount} 个文件 · {plan.rowCountSummary.plannedRowCount.toLocaleString()} 行 · {plan.columns.length} 列</strong>
            <span>{analysis.findingCount.toLocaleString()} 条数据异常候选，仅供核对；没有自动转换或修改。</span>
            {unsafeCells > 0 && <span>发现 {unsafeCells} 个以公式符号开头的文字单元格，CSV 下载已暂停。</span>}
          </div>

          <div className="sheet-batch-sources">
            {plan.sources.map((source) => (
              <p key={source.sourceLabel}>
                <strong>{source.sourceLabel}</strong>：{source.rowCount.toLocaleString()} 行
                {source.missingColumns.length > 0 && `；缺少列 ${source.missingColumns.join('、')}`}
                {source.extraColumns.length > 0 && `；新增列 ${source.extraColumns.join('、')}`}
              </p>
            ))}
          </div>

          <div className="sheet-batch-table-wrap" role="region" aria-label="合并结果前八行预览" tabIndex={0}>
            <table>
              <thead><tr>{plan.columns.map((column) => <th key={column}>{column}</th>)}<th>来源与原始行</th></tr></thead>
              <tbody>
                {previewRows.map((row) => (
                  <tr key={`${row.sourceLabel}-${row.originalRowIndex}`}>
                {row.values.map((value, columnIndex) => (
                      <td key={columnIndex}>{value === undefined ? '—' : value === null ? '空值' : String(value)}</td>
                    ))}
                    <td className="sheet-batch-provenance">{row.sourceLabel} · 第 {row.originalRowIndex + 1} 行</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="sheet-batch-provenance-note">预览显示前 8 行；下载会包含全部数据及来源文件／原始行号列。</p>

          <section className="sheet-batch-findings" aria-label="数据异常候选">
            <h3>数据检查候选（前 30 条）</h3>
            {!analysis.findingCount ? (
              <p>当前检查规则没有发现候选异常。</p>
            ) : (
              <ul>
                {previewFindings.map((finding, index) => (
                  <li key={`${finding.sourceLabel}-${finding.columnName}-${finding.rowIndex}-${finding.kind}-${index}`}>
                    <strong>{finding.sourceLabel} · {findingLocation(finding)}</strong>
                    <span>{finding.suggestion}</span>
                  </li>
                ))}
              </ul>
            )}
            {analysis.findingCount > previewFindings.length && (
              <p>共 {analysis.findingCount.toLocaleString()} 条，当前仅显示前 {previewFindings.length} 条。</p>
            )}
          </section>
        </>
      )}
    </section>
  );
}
