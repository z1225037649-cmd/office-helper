# 新增功能

以“查看选中字数”为例，使用已有的只读能力，因此无需修改 Office/WPS 适配器或 UI。

## 1. 新建功能模块

创建 `src/features/count-text/feature.ts`：

```ts
import type { FeatureDefinition } from '../../core/contracts';

export const countText: FeatureDefinition = {
  id: 'text.count',
  title: '数一数字数',
  description: '看看选中的文字有多长',
  icon: 'inspect',
  kinds: ['text'],
  requires: ['selection.read'],
  phrases: ['数一数字数', '多少字'],
  prepare(selection) {
    const count = Array.from(selection.text ?? '').length;
    return {
      summary: '统计当前选中的字符数量，包含空格和标点。',
      plan: { type: 'report', message: `当前选中 ${count} 个字符。` },
    };
  },
};
```

## 2. 注册

在 `src/features/index.ts` 导入 `countText`，加到 `features` 数组。界面自动产生功能卡片，本地指令匹配自动使用 `phrases`。如果需要功能开关，在 `src/app/config.ts` 添加 `'text.count': true`，改成 `false` 同时关闭入口与执行权限。

## 3. 检查

```bash
npm run check
npm run dev
```

在演示文档中操作。对于算法或写入相关功能，增加覆盖实际行为的单元测试；无需为纯文案或图标添加镜像测试。

## 新能力与 AI

如果新功能需要现有宿主契约没有的能力，例如设置段落对齐：

1. 在 `core/contracts.ts` 增加能力标识及有限的 `ActionPlan` 类型。
2. 在 `TaskRunner.validatePlan` 校验参数。
3. 在相应适配器实现，实机验证通过后再声明能力。
4. 定义独立功能并注册。暂未实现的宿主保持不声明此能力。

接入 AI 时实现 `IntentProvider` 并通过 `createRuntime(host, provider)` 注入。目前接口返回功能 ID，无参数；需要参数时先扩展结构化参数契约与运行时校验。生产密钥应由后端管理，不写入 `VITE_*` 环境变量或浏览器包；不要把用户文档内容当作系统指令。未来的语音模块只需把转录文本交给这个入口。
