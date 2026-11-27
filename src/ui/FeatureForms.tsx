import { useState } from 'react';
import type {
  DuplicateOptions,
  ExportScope,
  FeatureDefinition,
  FeatureParameters,
  NoticeFields,
  SelectionSnapshot,
} from '../core/contracts';

export interface FeatureFormContext {
  feature: FeatureDefinition;
  selection: SelectionSnapshot;
}

export function FeatureForms({
  context,
  busy,
  initialParameters,
  onSubmit,
  onCancel,
}: {
  context: FeatureFormContext;
  busy: boolean;
  initialParameters?: FeatureParameters;
  onSubmit: (parameters: FeatureParameters) => void;
  onCancel: () => void;
}) {
  const { feature, selection } = context;
  const [duplicates, setDuplicates] = useState<DuplicateOptions>(
    initialParameters?.duplicates ?? {
      keyColumns: [],
      requiredColumns: [],
      trimWhitespace: true,
      ignoreCase: false,
      hasHeader: false,
    },
  );
  const [notice, setNotice] = useState<NoticeFields>(
    initialParameters?.notice ?? { title: '', time: '', location: '', body: '', contact: '' },
  );
  const baseName = selection.documentName.replace(/\.[^.]+$/, '');
  const extension =
    feature.id === 'document.exportPdf'
      ? '.pdf'
      : (selection.documentName.match(/\.[a-z0-9]{2,5}$/iu)?.[0] ??
        (selection.kind === 'sheet' ? '.xlsx' : '.docx'));
  const [fileName, setFileName] = useState(
    initialParameters?.fileName ?? `${baseName}_副本${extension}`,
  );
  const [folder, setFolder] = useState(initialParameters?.folder ?? '');
  const [exportScope, setExportScope] = useState<ExportScope>(
    initialParameters?.exportScope ?? 'document',
  );
  const columnCount = Math.max(0, ...(selection.values ?? []).map((row) => row.length));
  const noticeLabels: { key: keyof NoticeFields; label: string; max: number }[] = [
    { key: 'title', label: '通知标题', max: 120 },
    { key: 'time', label: '时间', max: 200 },
    { key: 'location', label: '地点', max: 300 },
    { key: 'body', label: '正文', max: 12000 },
    { key: 'contact', label: '联系人及联系方式', max: 300 },
  ];

  function toggleColumn(key: 'keyColumns' | 'requiredColumns', column: number) {
    setDuplicates((value) => ({
      ...value,
      [key]: value[key].includes(column)
        ? value[key].filter((item) => item !== column)
        : [...value[key], column].sort((a, b) => a - b),
    }));
  }

  return (
    <form
      className="feature-form"
      aria-label={`${feature.title}设置`}
      onSubmit={(event) => {
        event.preventDefault();
        onSubmit(
          feature.form === 'duplicates'
            ? { duplicates }
            : feature.form === 'notice'
              ? { notice }
              : {
                  fileName,
                  ...(folder.trim() ? { folder: folder.trim() } : {}),
                  ...(feature.id === 'document.exportPdf' ? { exportScope } : {}),
                },
        );
      }}
    >
      <h2>{feature.title}</h2>
      <p className="form-target">
        {selection.documentName} · {selection.targetLabel}
      </p>
      {feature.form === 'duplicates' && (
        <>
          <p>请明确选择判重列或必填列。只列出疑似问题，供你核对。</p>
          <fieldset disabled={busy}>
            <legend>检查哪些列</legend>
            <div className="column-options">
              {Array.from({ length: columnCount }, (_, column) => (
                <div className="column-option" key={column}>
                  <span title={String(selection.values?.[0]?.[column] ?? '')}>
                    选区第 {column + 1} 列 ·{' '}
                    {String(selection.values?.[0]?.[column] ?? '空白').slice(0, 30)}
                  </span>
                  <label>
                    <input
                      type="checkbox"
                      aria-label={`选区第 ${column + 1} 列用于判重`}
                      checked={duplicates.keyColumns.includes(column)}
                      onChange={() => toggleColumn('keyColumns', column)}
                    />
                    判重
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      aria-label={`选区第 ${column + 1} 列为必填`}
                      checked={duplicates.requiredColumns.includes(column)}
                      onChange={() => toggleColumn('requiredColumns', column)}
                    />
                    必填
                  </label>
                </div>
              ))}
            </div>
            <label className="check-option">
              <input
                type="checkbox"
                checked={duplicates.hasHeader}
                onChange={(event) =>
                  setDuplicates({ ...duplicates, hasHeader: event.target.checked })
                }
              />
              第一行是表头，跳过检查
            </label>
            <label className="check-option">
              <input
                type="checkbox"
                checked={duplicates.trimWhitespace}
                onChange={(event) =>
                  setDuplicates({ ...duplicates, trimWhitespace: event.target.checked })
                }
              />
              忽略文字首尾空格
            </label>
            <label className="check-option">
              <input
                type="checkbox"
                checked={duplicates.ignoreCase}
                onChange={(event) =>
                  setDuplicates({ ...duplicates, ignoreCase: event.target.checked })
                }
              />
              文字判重忽略大小写
            </label>
          </fieldset>
        </>
      )}
      {feature.form === 'notice' && (
        <>
          <p>使用固定通知模板，以下五项都需要填写。确认后创建新文档。</p>
          {noticeLabels.map(({ key, label, max }) => (
            <label className="form-field" key={key}>
              {label}
              {key === 'body' ? (
                <textarea
                  required
                  disabled={busy}
                  value={notice[key]}
                  maxLength={max}
                  rows={6}
                  onChange={(event) => setNotice({ ...notice, [key]: event.target.value })}
                />
              ) : (
                <input
                  required
                  disabled={busy}
                  value={notice[key]}
                  maxLength={max}
                  onChange={(event) => setNotice({ ...notice, [key]: event.target.value })}
                />
              )}
            </label>
          ))}
        </>
      )}
      {feature.form === 'save' && (
        <>
          {feature.id === 'document.exportPdf' ? (
            <label className="form-field">
              导出范围
              <select
                disabled={busy}
                value={exportScope}
                onChange={(event) => setExportScope(event.target.value as ExportScope)}
              >
                <option value="document">
                  {selection.kind === 'sheet' ? '整个工作簿，含全部工作表' : '整篇文档'}
                </option>
                <option value="selection">仅当前选区</option>
                {selection.kind === 'sheet' && <option value="activeSheet">仅当前工作表</option>}
              </select>
            </label>
          ) : (
            <p>
              副本范围：{selection.kind === 'sheet' ? '整个工作簿，含全部工作表' : '整篇文档'}。
            </p>
          )}
          <p>使用新名称，保留原稿；重名时请更换名称。原生保存窗口仍需由你确认。</p>
          <label className="form-field">
            新文件名
            <input
              required
              disabled={busy}
              maxLength={180}
              value={fileName}
              onChange={(event) => setFileName(event.target.value)}
            />
          </label>
          <label className="form-field">
            目标文件夹（可选）
            <input
              disabled={busy}
              maxLength={1000}
              placeholder="不填则在保存窗口选择"
              value={folder}
              onChange={(event) => setFolder(event.target.value)}
            />
          </label>
          <p className="form-note">
            宿主不能直接选择文件夹或核验落盘时，会说明原生操作步骤或浏览器下载位置的核对方法。
          </p>
        </>
      )}
      <div className="preview-buttons">
        <button
          className="primary-button"
          type="submit"
          disabled={
            busy ||
            (feature.form === 'duplicates' &&
              !duplicates.keyColumns.length &&
              !duplicates.requiredColumns.length)
          }
        >
          预览下一步
        </button>
        <button className="secondary-button" type="button" disabled={busy} onClick={onCancel}>
          取消
        </button>
      </div>
    </form>
  );
}
