# 宿主接入

这是 Windows 开发机接入说明。Microsoft Word、Excel、WPS 文字和 WPS 表格已完成真实任务窗格的最小验证。成品安装包与生产发布尚未完成，版本及结果见 [验证记录](./verification.md)。

## Microsoft Office

1. 安装依赖：`npm ci`。
2. 本机已获用户授权，通过微软 `office-addin-dev-certs` 安装 30 天开发证书。证书文件位于当前用户的 `.office-addin-dev-certs` 目录；`npm run dev:office` 读取其中的 `localhost.key` 和 `localhost.crt`。在其他开发机上需先完成该机的开发证书安装与信任。
3. 运行 `npm run setup:office` 注册 `manifests/office/manifest.xml`，并生成 `output/samples/OfficeHelper-Sample.docx`、`OfficeHelper-Sample.xlsx`。已有同名试用文件会保留，避免覆盖试用内容。
4. 运行 `npm run dev:office`，保持 https://localhost:3443 服务运行。
5. 用 Word／Excel 打开相应试用文件，从加载项入口使用侧边栏。Word 先选中文字，再预览并确认字号修改；Excel 选中金额区域查看合计。

本机 WPS 可能接管 `.docx` 文件关联和 `Word.Application` COM 名称；双击文件不保证打开 Microsoft Word。可运行 `powershell.exe -NoProfile -File scripts/open-office-samples.ps1 -App Word`（表格用 `-App Excel`），该脚本明确指定常见安装路径下的微软程序；其他安装路径请从对应软件内打开文件。

清单权限为 `ReadWriteDocument`。WordApi 1.1 支持读取，WordApi 1.3 才开放字号修改与恢复；ExcelApi 1.1 支持读取。启动时检测能力，不支持时明确提示。Office.js 仅在 Office 入口从微软官方 CDN 加载。

- `npm run validate:office`：校验开发清单；校验结果不等于商店审核通过。
- `npm run remove:office`：移除本项目的 Office 开发注册。
- `npx office-addin-dev-certs uninstall`：卸载开发证书；与本项目注册分开处理。若其他开发加载项也依赖该证书，卸载后它们也可能需要重新配置。

开发证书和 localhost 地址不能直接用于生产分发。正式发布需受信任的 HTTPS 服务、相应清单 URL 和目标平台分发流程。

## WPS

1. 运行 `npm run setup:wps` 生成文字／表格入口并注册开发加载项。
2. 运行 `npm run dev`，保持 http://127.0.0.1:5173 服务运行。
3. 打开 WPS，使用“小帮手 → 帮我一下”。若之前已运行 WPS，保存自己的工作后重新启动，以加载更新的注册。文字端先选中文字测试放大与恢复，表格端选中金额查看合计。

两个入口分别为 `/wps/writer/`、`/wps/sheet/`，共用侧边栏分别带 `?host=wps&kind=text`、`?host=wps&kind=sheet`。`CreateTaskPane` 返回任务窗格对象或句柄时，入口分别处理；若返回句柄，再调用 `GetTaskPane`。

注册脚本读取 `%APPDATA%/kingsoft/wps/jsaddons/publish.xml`，仅合并 `OfficeHelperWriter`、`OfficeHelperSheet` 两个节点，保留其他插件。已有配置先备份到项目 `.local/wps-registration-backups/`。`npm run remove:wps` 只移除这两个节点；不会删除其他插件或试用文件。配置结构不符合预期时脚本停止。

本机 WPS 12.1.0.19770 的文字与表格基础链路通过，不代表所有个人版、企业版或发布渠道均兼容。`manifests/wps/publish.example.xml` 是参考结构，不能覆盖用户已有配置。异步 JSAPI 等其他运行时应在适配层单独支持，不扩散到功能模块。

## 演示与实机验证

浏览器直接打开开发服务首页使用内存演示文件；真实用户从办公软件内打开助手。显式请求 Office／WPS 宿主但连接失败时显示错误，不回退演示。

`output/samples/` 是正常试用文件；`output/native/OfficeHelper-Native-Test.docx`、`.xlsx` 和 `native-check.html` 是开发者诊断专用。`npm run prepare:native` 生成独立测试清单和测试文件；报告通道需要本地测试令牌，结果写入 `output/native/results.jsonl`，不用于日常产品流程。

重跑原生测试时，先关闭上次的专用测试文件并停止开发服务，然后生成测试入口，使用 `npx office-addin-dev-settings register .local/office-native-check.xml` 与 `powershell.exe -NoProfile -File scripts/register-wps.ps1 -Action Test` 切换开发注册。在两个服务终端分别设置 `$env:HELPER_NATIVE_TEST_TOKEN = (Get-Content .local/native-token -Raw).Trim()`，再启动 Office／WPS 服务。用 `scripts/test-office-native.ps1` 或 `scripts/test-wps-native.ps1` 的 `-App Word`／`-App Excel` 准备固定选区；Office 脚本若发现 COM 指向 WPS 会停止，需要先明确启动微软程序。测试页只有文件名和内容标记均匹配时才修改字号，并立即恢复。

测试结束运行 `npm run setup:office`、`npm run setup:wps` 恢复正常入口；移除服务终端的 `HELPER_NATIVE_TEST_TOKEN` 环境变量并重启服务。生产构建不包含诊断页面。

## 后续真实版本验收

- 安装／卸载保留其他插件，重启后能找到入口。
- 空选区、混合字号、只读与保护状态给出清楚提示。
- 单格、多格、公式结果、文本数字、错误值、隐藏／筛选行和超限范围。
- 多文档、未保存文件和多窗口切换；预览后变化应拒绝旧任务。
- 字号修改与恢复、后续编辑冲突、原生撤销交互和保存后的恢复。
- 连接失败、Office.js 加载失败、接口缺失和开发证书失效。
- 320–480px 侧边栏、键盘操作、系统缩放及目标用户独立完成任务。

## 官方资料

- [Office 版本与能力集](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [开发加载方式](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [WPS 加载项开发与发布](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/wps-addin-development-instructions)
- [WPS 任务窗格](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/TaskPane/task-pane-overview)
