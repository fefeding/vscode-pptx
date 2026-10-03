// Webview 样式（作为字符串注入，避免额外资源请求）
export const CSS = `
:root {
  --bg: #1e1e1e;
  --panel: #252526;
  --panel-2: #2d2d30;
  --border: #3c3c3c;
  --text: #e6e6e6;
  --muted: #9d9d9d;
  --accent: #f9ab00;
  --accent-2: #4285f4;
  --danger: #f14c4c;
  --ok: #4ec9b0;
}
* { box-sizing: border-box; }
html, body { margin: 0; height: 100%; }
body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
  background: var(--bg);
  color: var(--text);
  font-size: 13px;
  overflow: hidden;
}
#app { display: flex; flex-direction: column; height: 100vh; }

/* 预览模式下隐藏编辑相关控件 */
.app.mode-preview .edit-only { display: none !important; }

/* 顶栏 */
.topbar {
  display: flex; align-items: center; gap: 8px;
  padding: 6px 10px; background: var(--panel); border-bottom: 1px solid var(--border);
  flex: 0 0 auto;
}
.topbar .title { font-weight: 600; margin-right: auto; color: var(--accent); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 40vw; }
.btn {
  background: var(--panel-2); color: var(--text); border: 1px solid var(--border);
  border-radius: 4px; padding: 4px 10px; cursor: pointer; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;
}
.btn:hover { background: #3a3a3d; }
.btn.primary { background: var(--accent-2); border-color: var(--accent-2); color: #fff; }
.btn.primary:hover { filter: brightness(1.1); }
.btn:disabled { opacity: .45; cursor: default; }
.sep { width: 1px; height: 20px; background: var(--border); margin: 0 4px; }

/* 主区域 */
.main { display: flex; flex: 1 1 auto; min-height: 0; }
.slides-panel { width: 180px; background: var(--panel); border-right: 1px solid var(--border); display: flex; flex-direction: column; flex: 0 0 auto; }
.sp-head { padding: 6px 8px; font-size: 12px; color: var(--muted); display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid var(--border); }
.sp-list { overflow-y: auto; padding: 6px; display: flex; flex-direction: column; gap: 6px; }
.thumb {
  position: relative; border: 2px solid transparent; border-radius: 4px; cursor: pointer; background: #fff;
  overflow: hidden; flex: 0 0 auto;
}
.thumb.active { border-color: var(--accent); }
.thumb .num { position: absolute; top: 2px; left: 4px; font-size: 10px; color: #333; background: rgba(255,255,255,.7); border-radius: 3px; padding: 0 3px; }
.thumb .inner { transform-origin: top left; pointer-events: none; }

.canvas-area { flex: 1 1 auto; min-width: 0; display: flex; flex-direction: column; background: #161616; }
.canvas-scroll { flex: 1 1 auto; overflow: auto; display: flex; align-items: center; justify-content: center; padding: 20px; }
/* 幻灯片内部不应继承编辑器 UI 的字体设置（字号/字体族会改变文本度量与换行，
   与 examples/index.html 的默认值对齐：16px + 无衬线栈） */
.slide {
  font-size: 16px;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
}

/* 外层尺寸 = 缩放后尺寸（保证滚动范围与居中正确），内层按 1:1 布局后整体 scale */
.stage { position: relative; background: #fff; box-shadow: 0 4px 20px rgba(0,0,0,.4); overflow: hidden; flex: 0 0 auto; }
.stage-inner { position: absolute; top: 0; left: 0; transform-origin: top left; }
.slide-host { position: absolute; top: 0; left: 0; }
.slide-host .slide { margin: 0 !important; }
.overlay { position: absolute; top: 0; left: 0; pointer-events: none; }
.grid-overlay { position: absolute; inset: 0; pointer-events: none; display: none; }
.grid-overlay.on { display: block; background-image: linear-gradient(rgba(0,0,0,.08) 1px, transparent 1px), linear-gradient(90deg, rgba(0,0,0,.08) 1px, transparent 1px); background-size: 40px 40px; }
.sel-rect { position: absolute; border: 1px solid var(--accent); background: rgba(249,171,0,.08); pointer-events: auto; cursor: move; }
.sel-rect.selected { border-color: var(--accent-2); background: rgba(66,133,244,.1); }
.handle { position: absolute; width: 8px; height: 8px; background: #fff; border: 1px solid var(--accent-2); border-radius: 1px; }
.handle.nw { left: -4px; top: -4px; cursor: nwse-resize; }
.handle.n  { left: 50%; top: -4px; margin-left: -4px; cursor: ns-resize; }
.handle.ne { right: -4px; top: -4px; cursor: nesw-resize; }
.handle.e  { right: -4px; top: 50%; margin-top: -4px; cursor: ew-resize; }
.handle.se { right: -4px; bottom: -4px; cursor: nwse-resize; }
.handle.s  { left: 50%; bottom: -4px; margin-left: -4px; cursor: ns-resize; }
.handle.sw { left: -4px; bottom: -4px; cursor: nesw-resize; }
.handle.w  { left: -4px; top: 50%; margin-top: -4px; cursor: ew-resize; }

/* 右侧属性面板 */
.inspector { width: 280px; background: var(--panel); border-left: 1px solid var(--border); overflow-y: auto; padding: 10px; flex: 0 0 auto; }
.inspector h3 { font-size: 12px; margin: 0 0 8px; color: var(--accent); text-transform: uppercase; letter-spacing: .04em; }
.inspector h4 { font-size: 11px; margin: 12px 0 6px; color: var(--muted); }
.field { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; }
.field label { width: 64px; color: var(--muted); flex: 0 0 auto; }
.field input[type=text], .field input[type=number], .field select, .field textarea {
  flex: 1 1 auto; width: 100%; background: var(--panel-2); border: 1px solid var(--border); color: var(--text);
  border-radius: 4px; padding: 4px 6px; font-size: 12px; font-family: inherit;
}
.field input[type=color] { width: 36px; height: 26px; padding: 0; border: 1px solid var(--border); background: none; }
.field textarea { resize: vertical; min-height: 60px; }
.row { display: flex; gap: 6px; flex-wrap: wrap; }
.row .btn { flex: 1 1 auto; justify-content: center; }
.chk { display: flex; align-items: center; gap: 6px; }
.el-list { list-style: none; margin: 0; padding: 0; }
.el-list li { padding: 5px 8px; border: 1px solid var(--border); border-radius: 4px; margin-bottom: 4px; cursor: pointer; display: flex; justify-content: space-between; gap: 6px; }
.el-list li.active { border-color: var(--accent-2); background: var(--panel-2); }
.el-list li .tag { color: var(--muted); font-size: 11px; }
.empty { color: var(--muted); font-size: 12px; padding: 8px 0; }

/* 状态栏 */
.statusbar { display: flex; align-items: center; gap: 8px; padding: 4px 10px; background: var(--panel); border-top: 1px solid var(--border); font-size: 12px; color: var(--muted); flex: 0 0 auto; }
.statusbar .spacer { flex: 1 1 auto; }
.chip { background: var(--panel-2); border: 1px solid var(--border); border-radius: 4px; padding: 2px 8px; cursor: pointer; }
.chip:hover { background: #3a3a3d; }

/* 演示模式 */
.present { position: fixed; inset: 0; background: #000; display: none; align-items: center; justify-content: center; z-index: 50; }
.present.on { display: flex; }
.present .stage { box-shadow: none; background: #000; }
.present .slide-host { position: relative; transform-origin: center center; }
.present .slide-host .slide { margin: 0 !important; }

/* 文档属性面板（预览与编辑模式均可打开） */
.doc-info {
  position: fixed; top: 40px; right: 12px; width: 320px; max-height: 70vh; overflow: auto;
  background: var(--panel); border: 1px solid var(--border); border-radius: 6px; padding: 10px;
  display: none; z-index: 55; box-shadow: 0 6px 24px rgba(0,0,0,.45);
}
.doc-info.on { display: block; }
.doc-info-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; font-weight: 600; color: var(--accent); }
.doc-info h4 { font-size: 11px; margin: 10px 0 4px; color: var(--muted); }
.doc-info table.kv { width: 100%; border-collapse: collapse; font-size: 12px; }
.doc-info table.kv td { padding: 3px 4px; border-bottom: 1px solid var(--border); vertical-align: top; word-break: break-all; }
.doc-info table.kv td.k { width: 88px; color: var(--muted); white-space: nowrap; }

/* Toast */
.toast { position: fixed; bottom: 48px; left: 50%; transform: translateX(-50%); background: #333; color: #fff; padding: 6px 14px; border-radius: 6px; font-size: 12px; opacity: 0; transition: opacity .2s; pointer-events: none; z-index: 60; }
.toast.show { opacity: 1; }
.toast.error { background: var(--danger); }
.toast.warn { background: #b89500; }

.muted { color: var(--muted); }
`;
