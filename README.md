# PPTX Editor for VSCode

基于 [`@fefeding/ppt-parser`](https://github.com/fefeding/pptx-parser) 的 VSCode 插件，支持**默认以可视化方式打开并编辑 `.pptx` 文件**。

## 功能

- **默认打开 pptx**：在资源管理器中双击 `.pptx` 即以其内置编辑器打开（无需另装插件）。
- **幻灯片预览**：按原始比例渲染每一页，左侧缩略图导航，支持缩放、适应窗口、网格。
- **可视化编辑**：
  - 点击画布上的元素进行选择，拖拽移动、拖拽控制点缩放。
  - 属性面板编辑：位置/尺寸/旋转、文本（内容/字号/颜色/对齐/粗体/斜体/下划线）、形状（填充/边框）、图片替换。
  - 幻灯片操作：新增空白页、复制、删除、上移/下移、设置背景色、隐藏、演讲者备注。
  - 新增元素：文本框、矩形、图片（本地文件）。
- **演示模式**：F5 / 点击「演示」全屏放映，方向键或点击翻页。
- **保存与导出**：`Ctrl/Cmd+S` 保存；「导出」另存为副本；顶栏「放弃更改并重新加载」。
- **撤销 / 重做**：`Ctrl/Cmd+Z`、`Ctrl/Cmd+Y`。

## 开发

```bash
npm install            # 会安装 pptx-parser（file:../pptx-parser）及其依赖
npm run compile        # 用 esbuild 打包 dist/extension.js 与 dist/webview.js
npm run watch          # 监听模式
```

在 VSCode 中按 `F5` 以扩展开发宿主运行，打开任意 `.pptx` 即可。

## 说明与限制

- 编辑采用「标准 JSON 模型（PptxDocument）↔ PPTX」双向往返（由 pptx-parser 提供）。
  因此保存时会重新生成演示文稿，原文件中的**母版 / 版式**会按单一空白版式重写；
  文本框、形状、图片、图表、表格等**内容元素**可正常编辑与保留。
- 对 SmartArt / OLE / 公式等复杂对象以原始 OOXML（`__raw`）兜底保留，但暂不支持可视化编辑。
