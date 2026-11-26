# GitHub 参考项目

检查日期：2026-11-26。以下为源码结构研究，不代表其功能已经在我们的产品或目标用户电脑上验证。检查仓库 README、相关源码、许可证和 GitHub 元数据后，独立编写本工程；未运行任何参考仓库的安装脚本。

| 项目                                                                                              | 借鉴的设计                                                                              | 本工程的取舍                                                                                     |
| ------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| [OfficeDev/Office-Addin-TaskPane-React](https://github.com/OfficeDev/Office-Addin-TaskPane-React) | React / TypeScript 任务窗格、独立 manifest、宿主初始化                                  | 用 Vite 建立轻量开发工程；保留独立 Office 清单，功能不放在界面组件里                             |
| [Sivan22/autoOffice](https://github.com/Sivan22/autoOffice)                                       | `src/taskpane/host/context.ts` 集中识别不同 Office 宿主                                 | 采用统一宿主契约；真实宿主连接失败会报错，不默认冒充 Word；不采用动态生成并执行 Office.js 的模式 |
| [lewis-hui1202/WPS-AI](https://github.com/lewis-hui1202/WPS-AI)                                   | `plugin/js/hosts/*` 宿主封装、`plugin/js/tools/registry.js` 工具注册、独立 WPS 加载入口 | 借鉴宿主和功能分离；当前先用确定性操作计划，保留 AI 服务扩展口                                   |
| [humoumou1215/wps-assistant](https://github.com/humoumou1215/wps-assistant)                       | 加载项与本地服务分离，可包装桌面分发                                                    | 作为后续安装/桥接方案参考；初版不额外引入本地服务、MCP 或 Rust                                   |

微软模板的 GitHub 元数据返回 `NOASSERTION`，但已直接读取其 `LICENSE`，内容为 MIT。AutoOffice、WPS-AI、wps-assistant 的 GitHub 许可元数据为 MIT；正式复用文件前仍应固定提交并核对该文件及所带资源的许可。此次没有复制其业务源码，因此没有导入上游依赖或安装器。

检索还发现 [SourceFlowDigital/sourceflower-wps-assistant](https://github.com/SourceFlowDigital/sourceflower-wps-assistant) 与 [blazzbyte/agent-office](https://github.com/blazzbyte/agent-office)。前者聚焦 WPS 文字 AI 功能，后者采用 VSTO/COM 和 WebView2。它们说明侧边栏助手有多种实现方式；本工程先选共用 Web UI + 宿主适配层，减少双平台功能重复。

面向本产品的核心差异是：主要入口为简单任务按钮，界面明确处理范围、操作前预览、能力未接入时不显示成功；模型和配置细节不放进日常操作流程。
