import { FileText, Table2, ChevronDown, Check } from 'lucide-react';
import type { DocumentKind, SelectionSnapshot } from '../core/contracts';

interface Props {
  selection?: SelectionSnapshot;
  kind: DocumentKind;
  disabled: boolean;
  onSwitch: (kind: DocumentKind) => void;
}

export function DemoDocument({ selection, kind, disabled, onSwitch }: Props) {
  return (
    <section className="demo-workspace" aria-label="演示文档">
      <div className="workspace-toolbar">
        <div className="document-switch" role="group" aria-label="选择演示文件">
          <button
            disabled={disabled}
            className={kind === 'text' ? 'active' : ''}
            aria-pressed={kind === 'text'}
            onClick={() => onSwitch('text')}
          >
            <FileText size={17} />
            文字文档
          </button>
          <button
            disabled={disabled}
            className={kind === 'sheet' ? 'active' : ''}
            aria-pressed={kind === 'sheet'}
            onClick={() => onSwitch('sheet')}
          >
            <Table2 size={17} />
            电子表格
          </button>
        </div>
        <span className="workspace-hint">
          试一试右边的小帮手 <span aria-hidden="true">↗</span>
        </span>
      </div>
      <div className="document-window">
        <div className="window-title">
          <span className={`file-icon ${kind}`}>
            <FileText size={16} />
          </span>
          {selection?.documentName}
          <span className="sample-tag">示例</span>
        </div>
        <div className="ribbon" aria-hidden="true">
          <span>开始</span>
          <span>插入</span>
          <span>页面布局</span>
          <span className="ribbon-selected">小帮手</span>
        </div>
        <div className="format-bar" aria-hidden="true">
          <span>
            微软雅黑 <ChevronDown size={12} />
          </span>
          <span>
            {kind === 'text' ? selection?.fontSize : '12'} <ChevronDown size={12} />
          </span>
          <b>B</b>
          <i>I</i>
          <u>U</u>
          <span className="format-divider" />
          <span>≡</span>
          <span>☷</span>
        </div>
        {kind === 'text' ? (
          <div className="paper-surround">
            <article className="paper">
              <div className="paper-eyebrow">邻里相聚 · 美好生活</div>
              <h2
                style={{
                  fontSize: `${selection?.fontSize ?? 18}pt`,
                  textAlign:
                    selection?.paragraphs?.[0]?.alignment === 'Centered' ? 'center' : 'left',
                  lineHeight: (selection?.paragraphs?.[0]?.lineSpacing ?? 12) / 12,
                  textIndent: `${selection?.paragraphs?.[0]?.firstLineIndent ?? 0}pt`,
                }}
              >
                <mark
                  data-testid="demo-title"
                  style={{
                    fontSize: `${selection?.fontSize ?? 18}pt`,
                    fontWeight: selection?.bold ? 700 : 400,
                  }}
                >
                  社区活动通知
                </mark>
              </h2>
              <p>各位居民朋友：</p>
              <p>为了丰富大家的日常生活，社区将举办一场邻里交流活动，欢迎大家一起参加。</p>
              <div className="notice-details">
                <p>
                  <strong>活动时间</strong>
                  <span>本周六 上午 9:00</span>
                </p>
                <p>
                  <strong>活动地点</strong>
                  <span>社区活动中心</span>
                </p>
              </div>
              <p>带上一份好心情，认识身边的新朋友。期待与您相见！</p>
              <div className="paper-signature">社区服务中心</div>
              <div className="selection-caption">
                <span />
                当前已选中：通知标题
              </div>
            </article>
          </div>
        ) : (
          <div className="sheet-surround">
            <div className="formula-bar">
              <span>B2:B5</span>
              <i>ƒx</i>
              <span>120</span>
            </div>
            <table className="demo-sheet">
              <caption className="sr-only">活动费用，已选中金额列 B2 到 B5</caption>
              <thead>
                <tr>
                  <th />
                  <th>A</th>
                  <th>B</th>
                  <th>C</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <th>1</th>
                  <td className="table-heading">项目</td>
                  <td className="table-heading">金额（元）</td>
                  <td className="table-heading">备注</td>
                </tr>
                {['活动物料', '茶点水果', '宣传打印', '场地布置'].map((name, index) => (
                  <tr key={name}>
                    <th>{index + 2}</th>
                    <td>{name}</td>
                    <td className={`selected-cell cell-${index}`}>
                      {selection?.values?.[index]?.[0]}
                    </td>
                    <td>已登记</td>
                  </tr>
                ))}
                {[6, 7, 8, 9, 10].map((row) => (
                  <tr key={row}>
                    <th>{row}</th>
                    <td />
                    <td />
                    <td />
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="sheet-tab">
              <Table2 size={14} />
              费用表
            </div>
          </div>
        )}
        <div className="document-status">
          <span>
            <Check size={13} />
            演示内容
          </span>
          <span>{kind === 'text' ? '第 1 页，共 1 页' : '已选中 4 个单元格'}</span>
          <span>100%</span>
        </div>
      </div>
      <p className="demo-footnote">这是练习用的文件，可以放心尝试。刷新页面后恢复初始内容。</p>
    </section>
  );
}
