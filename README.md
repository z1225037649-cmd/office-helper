# 办公小帮手

面向不熟悉电脑的中老年人，在 WPS／Microsoft Office 内提供大按钮、明确的处理范围和操作预览。技术栈为 React + TypeScript + Vite，功能模块与办公软件适配器分开，方便逐个增加功能。

## 启动

需要 Node.js 22.12+（推荐 Node 24 LTS）。

```bash
npm ci
npm run dev
```

打开 http://127.0.0.1:5173 可进入**浏览器演示模式**，使用内存中的示例，不读写用户的 Office/WPS 文件，刷新会重置。这个服务同时提供 WPS 开发加载项页面；真实使用入口在办公软件内部，见下方接入步骤。

```bash
npm run check    # 类型检查、单元测试和生产构建
npm run build
npm run preview
```

## 当前完成程度

| 部分                                       | 状态                                              |
| ------------------------------------------ | ------------------------------------------------- |
| 共用侧边栏、文字/表格演示                  | 可运行                                            |
| 功能注册、功能开关、能力检查               | 已实现                                            |
| 操作预览、确认、过期选区检查、重复执行保护 | 已实现                                            |
| 演示文件修改、单步恢复、恢复冲突检查       | 已实现                                            |
| 本地短句匹配                               | 已实现，仅匹配已注册功能，不是大模型              |
| Microsoft Word 原生字号修改与恢复          | 真实任务窗格验证 18 → 22 → 18 磅通过              |
| Microsoft Excel 原生读取与合计             | 真实任务窗格读取四格，合计 450.5 通过             |
| WPS 表格原生读取与合计                     | 12.1.0.19770 真实任务窗格合计 450.5 通过          |
| WPS 文字原生字号修改与恢复                 | 12.1.0.19770 真实任务窗格验证 18 → 22 → 18 磅通过 |
| 语音、大模型、自动保存、安装包、生产部署   | 未实现，后续按需扩展                              |

Office 实测版本为 Windows x64 16.0.20625.20108；本轮 48 项单元测试、类型检查与生产构建通过。这些结果只覆盖本机测试场景，不代表所有版本和复杂文件均已通过；详见 [验证记录](docs/verification.md)。

合计只读取当前范围的数值类型单元格，结果只显示在助手内；文本数字、空白、布尔值和错误文本不计入并报告数量。隐藏／筛选行和不同宿主的错误值仍需专项验证。浮点计算最多显示 15 位有效数字。最多读取 5,000 个单元格或 20,000 字符。活动记录仅在内存中保留，不含正文；日常功能没有模型请求、遥测和密钥存储。

## 目录

```text
src/
  app/                配置与依赖装配
  core/               宿主契约、注册表、执行器、错误和选区校验
  features/           每个功能独立目录；index.ts 为注册入口
  adapters/
    mock/             可修改与恢复的演示宿主
    office/           Office.js 读取、Word 字号修改与恢复
    wps/              WPS JSAPI 读取、文字字号修改与恢复
    shared/           写入参数校验及默认只读基类
  services/           指令理解接口及本地实现
  ui/                 共用界面与演示文档
public/wps/           WPS 功能区入口与任务窗格回调
manifests/            Office 开发清单、WPS 发布配置示例
tests/unit/           核心执行约束与功能测试
src/adapters/**/*.test.ts  原生适配器边界回归测试
scripts/              开发注册、卸载、入口生成与实机验证辅助
docs/                架构、扩展步骤、宿主接入和开源参考
```

新增一个基于已有能力的功能，只需新建 `src/features/<功能>/feature.ts` 并在 `src/features/index.ts` 注册。界面和指令匹配自动获得入口；是否支持由功能声明和宿主能力共同决定。详见 [新增功能](docs/adding-features.md)。

- [架构与约束](docs/architecture.md)
- [宿主接入及验证清单](docs/host-integration.md)
- [功能路线图与五人试用计划](docs/product-roadmap.md)
- [GitHub 参考项目](docs/references.md)

## 开发用插件入口

- Office：已获授权安装微软官方开发证书，有效期 30 天。运行 `npm run setup:office` 注册开发加载项，再运行 `npm run dev:office` 提供 https://localhost:3443 页面；用 Word／Excel 打开 `output/samples/OfficeHelper-Sample.docx` 或 `.xlsx` 试用。`npm run validate:office` 校验清单，`npm run remove:office` 移除注册，`npx office-addin-dev-certs uninstall` 卸载开发证书。
- WPS：运行 `npm run setup:wps` 注册，再运行 `npm run dev` 提供 http://127.0.0.1:5173 页面，从“小帮手 → 帮我一下”打开。注册与 `npm run remove:wps` 只处理 `OfficeHelperWriter`、`OfficeHelperSheet` 两个节点；原有配置备份在 `.local/wps-registration-backups/`。
- `?host=office` 或 `?host=wps&kind=text` / `?host=wps&kind=sheet` 明确要求真实宿主。连接失败会显示错误，绝不自动切换为演示。

以上是开发机接入流程，依赖本地服务；成品安装包、自动更新和生产发布尚未完成。详细操作与限制见 [宿主接入](docs/host-integration.md)。

本项目独立编写，参考仓库仅用于结构研究，未整仓复制、运行其安装脚本或导入其业务代码。第三方 npm 依赖的许可证由对应包提供。本项目尚未决定对外发布许可证。
