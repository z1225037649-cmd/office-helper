// Development preparation only: no registration, certificate installation or host launch.
import { mkdir, readFile, writeFile, lstat, realpath } from 'node:fs/promises';
import { randomBytes, randomUUID, createHash } from 'node:crypto';
import { dirname, resolve, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import JSZip from 'jszip';

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const marker = 'OFFICE_HELPER_EXTENDED_TEST_V1';
export const matrices = ['office-text', 'office-sheet', 'wps-text', 'wps-sheet'];
export const rows = [
  ['编号', '姓名', '联系方式', '备注', '金额'],
  ['A-01', 'Alice', '1001', '常规', 10],
  ['A-02', ' alice ', '1001', '重复记录', 20],
  ['A-03', 'Bob', '', '漏填联系方式', 30],
  ['A-04', 'Carol', '1004', '', 0],
  ['A-05', 'Delta', '1005', '公式值：应保留公式而非只保留计算值', 30],
  ['A-06', 'ALICE', '1001', '隐藏行仍参与检查', 40],
  ['A-07', '', '1007', '漏填姓名，false不是空白', false],
];
const xml = (value) =>
  String(value).replace(
    /[&<>"']/gu,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c],
  );
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const textFor = (run, matrix) =>
  `${marker}\nRUN=${run}\nMATRIX=${matrix}\n仅供扩展能力隔离验收；原稿保持不变。`;
const newFile = (path, data) => writeFile(path, data, { flag: 'wx', mode: 0o600 });

function packageBase(main, contentType) {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${main}" ContentType="${contentType}"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${main}"/></Relationships>`,
  );
  return zip;
}

export async function wordFixture(config) {
  const zip = packageBase(
    'word/document.xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${config.text
      .split('\n')
      .map(
        (text) =>
          `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:p>`,
      )
      .join(
        '',
      )}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function sheetFixture(config) {
  const zip = packageBase(
    'xl/workbook.xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  );
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml').async('string')).replace(
      '</Types>',
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    ),
  );
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0"?><workbook xmlns="${ns}" xmlns:r="${relationship}"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="ExtendedData" sheetId="1" r:id="rId1"/><sheet name="ScopeSentinel" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">ExtendedData!$B$3:$D$6</definedName></definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationship}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${relationship}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${relationship}/styles" Target="styles.xml"/></Relationships>`,
  );
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0"?><styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf fontId="0" xfId="0"/><xf fontId="1" xfId="0" applyFont="1"/><xf fontId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  );
  const cell = (address, value, style = 0) =>
    `<c r="${address}" s="${style}"${typeof value === 'string' ? ' t="inlineStr"' : typeof value === 'boolean' ? ' t="b"' : ''}>${typeof value === 'string' ? `<is><t xml:space="preserve">${xml(value)}</t></is>` : `<v>${typeof value === 'boolean' ? Number(value) : value}</v>`}</c>`;
  const sheetRows =
    `<row r="1">${cell('A1', marker)}${cell('B1', config.run)}${cell('C1', config.matrix)}</row>` +
    rows
      .map(
        (row, i) =>
          `<row r="${i + 3}"${i === 6 ? ' hidden="1"' : ''}>${row.map((value, j) => (i === 5 && j === 4 ? '<c r="F8"><f>SUM(F4:F5)</f><v>30</v></c>' : cell(`${String.fromCharCode(66 + j)}${i + 3}`, value, i === 0 && j === 0 ? 1 : i === 1 && j === 3 ? 2 : 0))).join('')}</row>`,
      )
      .join('') +
    `<row r="12">${cell('B12', 'OFFICE_HELPER_EXTENDED_OUTSIDE')}${cell('F12', 909)}</row>`;
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${ns}"><sheetPr><pageSetUpPr fitToPage="0"/></sheetPr><dimension ref="A1:F12"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="B3" sqref="B3:F10"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="6" width="11" customWidth="1"/></cols><sheetData>${sheetRows}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" scale="85" fitToWidth="2" fitToHeight="3"/></worksheet>`,
  );
  zip.file(
    'xl/worksheets/sheet2.xml',
    `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${ns}"><dimension ref="A1:C1"/><sheetData><row r="1">${cell('A1', 'OFFICE_HELPER_EXTENDED_SECOND_SHEET')}${cell('B1', config.run)}${cell('C1', config.matrix)}</row></sheetData></worksheet>`,
  );
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export function wpsBootstrap(config, url, { registeredEntry = false } = {}) {
  // Only the registered development entry polls. Neither bridge ever runs a feature.
  return `(function(){
  var pane; var paneAttempted=false;
  var config=${JSON.stringify({ fileName: config.fileName, fixturePath: config.fixturePath, text: config.text, kind: config.kind, values: config.values, run: config.run, matrix: config.matrix })};
  function path(value){return String(value||'').split('/').join('\\\\').toLowerCase();}
  function stop(message){var node=document.getElementById('status');if(node)node.textContent=message;return false;}
  window.OnExtendedAddinLoad=function(){return true;};
  window.OnOpenExtendedCheck=function(){
    try{
      var api=window.wps;
      var app=window.Application||(api&&api.Application)||(api&&(config.kind==='text'?api.WpsApplication&&api.WpsApplication():api.EtApplication&&api.EtApplication()))||api;
      var doc=app&&(config.kind==='text'?app.ActiveDocument:app.ActiveWorkbook);
      if(!doc||doc.Name!==config.fileName||path(doc.FullName)!==path(config.fixturePath))return stop('拒绝：请打开本轮对应矩阵的专用文件。');
      if(config.kind==='text'){
        var text=doc.Content&&doc.Content.Text;
        if(typeof text!=='string'||text.replace(/\\r\\n?/g,'\\n').replace(/\\n$/,'')!==config.text)return stop('拒绝：完整正文标记不匹配。');
      }else{
        var sheet=app.ActiveSheet;
        if(!sheet||sheet.Name!=='ExtendedData'||typeof sheet.Range!=='function')return stop('拒绝：请选择 ExtendedData。');
        var mark=sheet.Range('A1:C1').Value2;
        if(JSON.stringify(mark)!==JSON.stringify([['${marker}',config.run,config.matrix]])||JSON.stringify(sheet.Range('B3:F10').Value2)!==JSON.stringify(config.values))return stop('拒绝：RUN/MATRIX/精确数据矩阵不匹配。');
      }
      var owner=typeof app.CreateTaskPane==='function'?app:api;
      if(!owner||typeof owner.CreateTaskPane!=='function')return stop('unsupported：没有真实任务窗格接口。');
      if(!pane){
        if(paneAttempted)return stop('unsupported：任务窗格已尝试创建，请人工复查，不重复创建。');
        paneAttempted=true;
        var made=owner.CreateTaskPane(${JSON.stringify(url)});
        pane=typeof made==='number'||typeof made==='string'?(typeof owner.GetTaskPane==='function'?owner.GetTaskPane(made):undefined):made;
      }
      if(!pane)return stop('unsupported：宿主未返回任务窗格。');
      pane.Visible=true;return true;
    }catch(error){return stop('unsupported：'+String(error));}
  };
  ${
    registeredEntry
      ? `window.OnAddinLoad=window.OnExtendedAddinLoad;
  window.OnOpenHelper=window.OnOpenExtendedCheck;
  setTimeout(function(){
    var count=0;
    var timer=setInterval(function(){
      count++; window.OnOpenHelper();
      if(paneAttempted||count>=90)clearInterval(timer);
    },700);
  },1200);`
      : ''
  }
})();`;
}

export function createWpsExtendedEntryMiddleware(configs) {
  const entries = new Map();
  for (const [entry, kind] of [
    ['writer', 'text'],
    ['sheet', 'sheet'],
  ]) {
    const matches = configs.filter(
      (config) => config.host === 'wps' && config.matrix === `wps-${kind}` && config.kind === kind,
    );
    if (matches.length !== 1) throw new Error('WPS 注册入口缺少唯一的专项矩阵。');
    const config = matches[0];
    entries.set(
      `/wps/${entry}/main.js`,
      wpsBootstrap(config, config.url, { registeredEntry: true }),
    );
  }
  return (request, response, next) => {
    const path = new URL(request.url ?? '/', 'http://127.0.0.1').pathname;
    const script = entries.get(path);
    if (!script) {
      next();
      return;
    }
    response.setHeader('Cache-Control', 'no-store');
    response.setHeader('Referrer-Policy', 'no-referrer');
    if (!['GET', 'HEAD'].includes(request.method)) {
      response.statusCode = 405;
      response.setHeader('Allow', 'GET, HEAD');
      response.end();
      return;
    }
    response.setHeader('Content-Type', 'application/javascript; charset=utf-8');
    response.end(request.method === 'HEAD' ? undefined : script);
  };
}

export async function prepareExtendedCheck(root = projectRoot, suppliedToken) {
  const token =
    suppliedToken ??
    process.env.HELPER_NATIVE_TEST_TOKEN ??
    (await readFile(join(root, '.local/native-token'), 'utf8')).trim();
  if (!/^[a-f0-9]{48}$/u.test(token))
    throw new Error('需要有效的本地 native token；本工具不改写已有 token。');
  const run = randomBytes(16).toString('hex');
  const parent = join(root, 'output/native/extended');
  await mkdir(parent, { recursive: true });
  const directory = join(parent, run);
  await mkdir(directory); // Never reuse an existing run, including an incomplete preparation.
  const template = await readFile(join(root, 'manifests/office/manifest.xml'), 'utf8');
  const configs = [];
  for (const matrix of matrices) {
    const [host, kind] = matrix.split('-');
    const fileName = `OfficeHelper-Extended-Test.${kind === 'text' ? 'docx' : 'xlsx'}`;
    const fixturePath = join(directory, 'fixtures', matrix, fileName);
    const resultDirectory = join(directory, 'results', matrix);
    await mkdir(dirname(fixturePath), { recursive: true });
    await mkdir(resultDirectory, { recursive: true });
    const config = {
      schema: 1,
      run,
      matrix,
      host,
      kind,
      fixturePath,
      resultDirectory,
      fileName,
      text: textFor(run, matrix),
      values: rows,
      notice: {
        title: `扩展诊断通知 ${run}`,
        time: '2026年11月28日 09:00（测试数据）',
        location: '隔离验收室（测试数据）',
        body: `${marker}\nRUN=${run}\nMATRIX=${matrix}`,
        contact: '测试联系人，不发送通知',
      },
    };
    const bytes = await (kind === 'text' ? wordFixture(config) : sheetFixture(config));
    await newFile(fixturePath, bytes);
    const base = host === 'office' ? 'https://localhost:3443' : 'http://127.0.0.1:5173';
    const url = `${base}/__extended/${run}/${matrix}/?host=${host}&kind=${kind}&token=${token}`;
    if (host === 'office') {
      const manifest = template
        .replace(/<Id>[^<]+<\/Id>/u, `<Id>${randomUUID()}</Id>`)
        .replace(/<Host Name="(?:Document|Workbook)"\s*\/>/gu, (match) =>
          match.includes(kind === 'text' ? 'Document' : 'Workbook') ? match : '',
        )
        .replace('https://localhost:3443/?host=office', xml(url))
        .replace('办公小帮手（开发版）', `扩展诊断 ${kind}（隔离开发）`);
      if (!manifest.includes(xml(url))) throw new Error('开发 manifest 入口不匹配。');
      await newFile(join(directory, `${matrix}.xml`), manifest);
    } else {
      const bootstrap = join(directory, 'wps', matrix);
      await mkdir(bootstrap, { recursive: true });
      await newFile(
        join(bootstrap, 'index.html'),
        `<!doctype html><html><meta charset="UTF-8"><body><p id="status">通过扩展专项工具栏按钮打开任务窗格；不会自动执行功能。</p><script src="main.js?token=${token}"></script></body></html>`,
      );
      await newFile(join(bootstrap, 'main.js'), wpsBootstrap(config, url));
      await newFile(
        join(bootstrap, 'ribbon.xml'),
        '<?xml version="1.0"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="OnExtendedAddinLoad"><ribbon><tabs><tab id="ExtendedCheckTab" label="隔离专项诊断"><group id="ExtendedCheckGroup" label="专用样本"><button id="OpenExtendedCheck" label="核验后打开诊断" size="large" onAction="OnOpenExtendedCheck"/></group></tab></tabs></ribbon></customUI>',
      );
      await newFile(
        join(bootstrap, 'manifest.xml'),
        `<?xml version="1.0"?><JsPlugin><ApiVersion>1.0.0</ApiVersion><Name>OfficeHelperExtended${kind === 'text' ? 'Writer' : 'Sheet'}${run}</Name><Description>本轮隔离开发诊断</Description></JsPlugin>`,
      );
    }
    configs.push({
      ...config,
      url,
      ...(host === 'wps'
        ? {
            bootstrapUrl: `${base}/__extended/${run}/${matrix}/bootstrap/index.html?token=${token}`,
          }
        : {}),
      sha256: hash(bytes),
    });
  }
  await newFile(join(directory, 'run.json'), JSON.stringify({ schema: 1, run, configs }, null, 2));
  await newFile(join(directory, 'CHECKLIST.md'), checklist(configs, directory));
  await newFile(
    join(directory, 'WPS-ENTRY.md'),
    `# WPS 专项入口（由主线程正常 UI 加载）\n\n先启动本轮 WPS 开发服务，勿运行旧 prepare-native-check 或覆盖旧 WPS main.js。\n\n\`node scripts/prepare-extended-check.mjs --serve "${directory}" --host wps\`\n\n已有正常注册 http://127.0.0.1:5173/wps/writer/ 与 /wps/sheet/ 时，无需改 publish.xml 或重新注册。本次诊断服务仅在内存中拦截两个路径的 main.js；HTML、ribbon 和磁盘 public 文件仍用原入口。保留 OnAddinLoad / OnOpenHelper 和约 64 秒的有界短轮询；仅当前文件 Name、FullName 完整路径、全部正文／精确矩阵及 RUN/MATRIX 与本轮配置完全匹配才创建对应诊断任务窗格。普通文件不创建诊断窗格。服务就绪前已缓存旧脚本的宿主需要由主线程通过正常 UI 重新加载入口；本工具不自动刷新宿主。\n\n打开 fixtures/wps-text/OfficeHelper-Extended-Test.docx 或 fixtures/wps-sheet/OfficeHelper-Extended-Test.xlsx。可等短轮询核验开窗，或点击原“小帮手 → 帮我一下”。入口仅打开对应任务窗格，不自动执行 feature。任务窗格再点“连接并核验专用样本”领取单元。Word/WPS文字全选四段正文；表格选择 ExtendedData!B3:F10。\n\nrun.json 中 bootstrapUrl 和本轮 wps/wps-text 或 wps/wps-sheet 仍是可选独立开发入口；仅在正常受支持的开发加载 UI 中人工使用。独立入口需显式点击“隔离专项诊断 → 核验后打开诊断”，不会自动轮询。已有获授权任务窗格支持手工输入 URL 时，也可使用对应 url。诊断总会重新核对路径、文件名、内容、公式、RUN/MATRIX 和原生适配器。URL 含本地 token，不粘到公共日志。不要在普通浏览器中把页面结果当作宿主验收。没有受支持入口则记 unsupported，不改注册表绕过。\n\n本轮禁止自动注册、启动应用、修改证书或已有专用样本。Excel 登录设置阻挡应记 blocked，等主线程通过正常界面解除后再验；尚未调用宿主接口不能称为 unsupported 或 passed。其余步骤见 CHECKLIST.md。\n`,
  );
  return { directory, run, configs, token };
}

function checklist(configs, directory) {
  return `# 扩展能力真实宿主验收准备（尚未实机执行）\n\n本轮：${configs[0].run}\n\n启动：\n\n\`node scripts/prepare-extended-check.mjs --serve "${directory}" --host office\`\n\n或 --host wps。仅复用现有 Vite 配置和 nativeReportPlugin；Office 只读取已存在的开发证书，缺失即失败，不安装。不要同时启动占用同端口的服务器。\n\n报告复用 output/native/results.jsonl；按 suite=extended、run、matrix 筛选。token 只在本轮 run.json 的本机 URL 内，勿分享。报告中的 claim 使用 extended:run:matrix；多任务窗格抢占即拒绝。服务重启后锁会重置，须重新生成一轮，不能重放旧目录。\n\n## 四个独立单元\n\n${configs.map((c) => `- ${c.matrix}: ${c.fixturePath}\n  - 结果唯一目录：${c.resultDirectory}\n  - URL 见 run.json；Office 使用本轮 ${c.matrix}.xml 通过正常开发加载流程人工加载。\n`).join('')}\nWPS 由已获授权的正常开发任务窗格入口人工打开 URL；本工具不写注册表、不生成或替换既有 WPS 启动脚本。若宿主不提供打开开发任务窗格的入口，记 unsupported，不能用浏览器替代。\n\n## 每个单元的操作\n\n1. 只打开该单元的专用文件；Word 全选四段正文，表格选择 ExtendedData!B3:F10。核验文件名、磁盘路径、RUN、MATRIX、全部正文／精确矩阵、F8=SUM(F4:F5)、第2表哨兵。隐藏的第9行不可删除或漏查。\n2. 每项先预览后取消；另测新预览替换旧预览（EXPIRED_TASK），以及预览后用原生 UI 移动选区，再点验证过期（STALE_SELECTION）。预览/取消必须保持正文、值、公式、顺序、格式和打印设置不变。切换文件测试只使用本轮另一个专用样本，人工记录，绝不在个人文件上尝试。\n3. 表格排版三项逐一确认并读回，再恢复并逐项比对原值。用原生 UI 修改本项属性后，验证恢复冲突；不得覆盖后续编辑。按页面显示的“确认后快照”手动还原冲突属性，再恢复到原始快照；恢复前不得开始下一项。\n4. 打印设置确认后核对选区区域、横向、宽1页/高自动；打开 Ctrl+P，用正常 UI 核对实际页数、漏列、可读性，取消返回；绝不提交打印。再恢复到原打印区域 B3:D6、纵向、85% 等实测原值。也测恢复冲突。缺接口显示 unsupported。\n5. 查重漏填预期 duplicate-2=C5、missing-3-2=D6、duplicate-6=C9（隐藏行）、missing-7-1=C10。逐一点击定位，确认原始行列；不删、排、写值。移动选区后仍定位原范围；改数据后定位应拒绝（人工记录，之后恢复专用样本）。\n6. 副本/PDF 只能用页面给出的唯一文件名和本单元新结果目录。Office 传入明确目录将走原生指引，自动接口计 unsupported；按正常 UI 补足保存对话框。在原生对话框测取消、已有路径时取消覆盖、选择其它目录时取消。已有路径测试只取消，不点击覆盖；不得写出专用目录。每种自动导出每轮只允许尝试一次，失败可能留残件，不得重试或覆盖。通过服务“只读核对结果文件”验证文件头、包结构和原样本磁盘哈希，再人工重新打开核对实际位置及整篇／所有工作表（含 ScopeSentinel）和 PDF 分页。正常 UI 手工完成单独记 manual，不把 unsupported/download/guidance 写成 passed。\n7. 通知先核对完整预览、取消和返回修改，再确认新建。核对原件未改、新窗口独立且未保存、五项字段完全一致；不发送、不保存原件。Office 无法从原任务窗格读取新窗口时计 manual-required；WPS 可读回时核验新文档身份和完整正文。保留新窗口供复查，手动返回专用原件后才进行下一项。\n\n## 证据规则\n\n页面只向同源 token 报告端点 POST；运行前必须成功 claim。failed 后停止，不自动修复或关闭宿主。指导、下载、取消输出和缺少接口均不算原生功能通过。清单人工项附实际软件版本、证据路径、原生对话框结果；人工项只能记 manual，不覆盖自动项状态。四单元均完成自动与人工证据前，不得宣称五项实机总验收通过。此目录不纳入生产包。\n`;
}

async function plainFile(path) {
  const stat = await lstat(path);
  if (!stat.isFile() || stat.isSymbolicLink()) throw new Error('拒绝链接或非普通文件。');
  return readFile(path);
}

export async function verifyOutput(config, feature) {
  if (!['document.saveCopy', 'document.exportPdf'].includes(feature))
    throw new Error('未知导出类型。');
  const extension =
    feature === 'document.exportPdf' ? 'pdf' : config.kind === 'text' ? 'docx' : 'xlsx';
  const fileName = `Extended-${config.run}-${config.matrix}.${extension}`;
  const path = join(config.resultDirectory, fileName);
  if (
    resolve(await realpath(config.resultDirectory)).toLowerCase() !==
    resolve(config.resultDirectory).toLowerCase()
  )
    throw new Error('结果目录不可是链接。');
  const source = await plainFile(config.fixturePath);
  if (hash(source) !== config.sha256) throw new Error('原始样本磁盘内容已变，拒绝报告导出通过。');
  const bytes = await plainFile(path);
  if (extension === 'pdf') {
    if (
      !bytes.subarray(0, 5).equals(Buffer.from('%PDF-')) ||
      !/%%EOF\s*$/u.test(bytes.subarray(-2048).toString('latin1'))
    )
      throw new Error('PDF 文件头或结束标记无效。');
  } else {
    const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
    const original = await JSZip.loadAsync(source);
    // The same native serialization may reorder XML; inspect part semantics manually too.
    const parts =
      extension === 'docx'
        ? ['word/document.xml']
        : ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'];
    for (const part of parts) {
      if (!zip.file(part) || !original.file(part)) throw new Error('副本缺少必要 OOXML 部件。');
    }
    const data = await zip.file(parts.at(-1)).async('string');
    if (!data.includes(config.run) || !data.includes(config.matrix))
      throw new Error('副本运行或矩阵标记不匹配。');
  }
  return {
    fileName,
    path,
    bytes: bytes.length,
    sha256: hash(bytes),
    status: 'manual-required',
    reason: '字节和原稿哈希已核对；请原生打开核对整篇/所有工作表、分页和实际位置。',
  };
}

export async function serveExtendedCheck(directory, host, root = projectRoot) {
  if (!['office', 'wps'].includes(host)) throw new Error('--host 必须为 office 或 wps。');
  const runDirectory = resolve(directory);
  const data = JSON.parse(await plainFile(join(runDirectory, 'run.json')));
  if (
    !/^[a-f0-9]{32}$/u.test(data.run) ||
    runDirectory !== join(root, 'output/native/extended', data.run)
  )
    throw new Error('不是项目内专用新诊断目录。');
  const configs = data.configs;
  if (
    !Array.isArray(configs) ||
    configs.length !== 4 ||
    configs.some(
      (c, i) =>
        c.matrix !== matrices[i] ||
        c.run !== data.run ||
        c.host !== c.matrix.split('-')[0] ||
        c.kind !== c.matrix.split('-')[1] ||
        c.fixturePath !== join(runDirectory, 'fixtures', c.matrix, c.fileName) ||
        c.resultDirectory !== join(runDirectory, 'results', c.matrix),
    )
  )
    throw new Error('运行配置不完整。');
  const token = new URL(configs[0].url).searchParams.get('token');
  if (
    !token ||
    !/^[a-f0-9]{48}$/u.test(token) ||
    (process.env.HELPER_NATIVE_TEST_TOKEN && process.env.HELPER_NATIVE_TEST_TOKEN !== token)
  )
    throw new Error('本地 token 不匹配。');
  for (const config of configs) {
    if (hash(await plainFile(config.fixturePath)) !== config.sha256)
      throw new Error('样本已被修改；请重新准备一轮。');
  }
  // Exclusive persistent lease: a restart must not replay an old export or claim.
  await newFile(join(runDirectory, `server-${host}.lock`), String(process.pid));
  process.env.HELPER_NATIVE_TEST_TOKEN = token;
  const reserved = new Set();
  const plugin = {
    name: 'isolated-extended-check-entry',
    configureServer(server) {
      if (host === 'wps') server.middlewares.use(createWpsExtendedEntryMiddleware(configs));
      server.middlewares.use(`/__extended/${data.run}`, (request, response) => {
        void (async () => {
          const url = new URL(request.url ?? '/', 'http://localhost');
          const [matrix, operation = '', bootstrapFile] = url.pathname.slice(1).split('/');
          const config = configs.find((c) => c.matrix === matrix && c.host === host);
          if (!config) {
            response.statusCode = 404;
            response.end();
            return;
          }
          if (operation === 'bootstrap' && config.host === 'wps' && request.method === 'GET') {
            if (
              !['index.html', 'main.js', 'manifest.xml', 'ribbon.xml'].includes(bootstrapFile) ||
              (['index.html', 'main.js'].includes(bootstrapFile) &&
                url.searchParams.get('token') !== token)
            ) {
              response.statusCode = 403;
              response.end();
              return;
            }
            response.setHeader(
              'Content-Type',
              bootstrapFile.endsWith('.js')
                ? 'application/javascript'
                : bootstrapFile.endsWith('.xml')
                  ? 'application/xml'
                  : 'text/html; charset=utf-8',
            );
            response.setHeader('Cache-Control', 'no-store');
            response.end(await plainFile(join(runDirectory, 'wps', matrix, bootstrapFile)));
            return;
          }
          if (!operation && request.method === 'GET') {
            response.setHeader('Content-Type', 'text/html; charset=utf-8');
            response.setHeader('Cache-Control', 'no-store');
            response.setHeader('Referrer-Policy', 'no-referrer');
            response.end(
              '<!doctype html><html lang="zh-CN"><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>扩展能力隔离诊断</title><body style="font:16px/1.6 system-ui;padding:16px;overflow-wrap:anywhere"><h1>扩展能力隔离诊断</h1><p>先核验专用样本。所有写入均需逐项预览确认；人工项及缺接口不能计作自动通过。</p><div id="extended-root"></div><script type="module" src="/src/dev/extended-check.ts"></script></body></html>',
            );
            return;
          }
          if (
            request.headers['x-helper-test'] !== token ||
            !['config', 'reserve-copy', 'reserve-pdf', 'verify-copy', 'verify-pdf'].includes(
              operation,
            ) ||
            request.method !== (operation === 'config' ? 'GET' : 'POST')
          ) {
            response.statusCode = 403;
            response.end();
            return;
          }
          response.setHeader('Content-Type', 'application/json; charset=utf-8');
          response.setHeader('Cache-Control', 'no-store');
          if (operation === 'config') {
            response.end(JSON.stringify(config));
            return;
          }
          const feature = operation.endsWith('pdf') ? 'document.exportPdf' : 'document.saveCopy';
          if (operation.startsWith('reserve')) {
            const key = `${matrix}:${feature}`;
            if (reserved.has(key)) throw new Error('本轮已尝试此导出；不能重试或覆盖。');
            const path = join(
              config.resultDirectory,
              `Extended-${config.run}-${matrix}.${feature === 'document.exportPdf' ? 'pdf' : config.kind === 'text' ? 'docx' : 'xlsx'}`,
            );
            try {
              await lstat(path);
              throw new Error('目标已存在，拒绝覆盖。');
            } catch (error) {
              if (error.code !== 'ENOENT') throw error;
            }
            if (
              resolve(await realpath(config.resultDirectory)).toLowerCase() !==
              resolve(config.resultDirectory).toLowerCase()
            )
              throw new Error('拒绝链接目录。');
            reserved.add(key);
            response.end(JSON.stringify({ reserved: true }));
          } else response.end(JSON.stringify(await verifyOutput(config, feature)));
        })().catch((error) => {
          response.statusCode = 409;
          response.end(JSON.stringify({ error: error.message }));
        });
      });
    },
  };
  const { createServer } = await import('vite');
  const server = await createServer({
    root,
    configFile: join(root, 'vite.config.ts'),
    mode: host === 'office' ? 'office' : 'development',
    plugins: [plugin],
    server: { host: host === 'office' ? 'localhost' : '127.0.0.1' },
  });
  await server.listen();
  console.log(
    `扩展诊断服务就绪（${host}）。入口和人工清单：${join(runDirectory, 'CHECKLIST.md')}。未启动任何宿主。`,
  );
  return server;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  if (process.argv[2] === '--serve') {
    if (!process.argv[3] || process.argv[4] !== '--host' || !process.argv[5])
      throw new Error('用法：--serve <本轮绝对目录> --host office|wps');
    await serveExtendedCheck(process.argv[3], process.argv[5]);
  } else {
    if (process.argv.length > 2) throw new Error('未知参数，未准备诊断。');
    const result = await prepareExtendedCheck();
    console.log(
      `已创建四单元隔离样本与清单：${result.directory}\n未连接真实宿主、注册加载项、安装证书或执行验收。`,
    );
  }
}
