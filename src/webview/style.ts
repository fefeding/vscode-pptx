// Webview styles (injected as a string to avoid extra resource requests)
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

/* Hide editing controls in preview mode */
.app.mode-preview .edit-only { display: none !important; }

/* Top toolbar (mirrors the editor's insert / arrange toolbar) */
.sp-topbar {
  display: flex; align-items: center; gap: 6px; flex-wrap: wrap;
  padding: 6px 10px; background: var(--panel); border-bottom: 1px solid var(--border);
}
.sp-topbar .btn { padding: 4px 10px; font-size: 12px; }
.tb-sep { width: 1px; height: 22px; background: var(--border); margin: 0 4px; }

/* Context menu (right-click on canvas) */
.ctx-menu {
  position: fixed; z-index: 70; min-width: 180px;
  background: var(--panel); border: 1px solid var(--border); border-radius: 6px;
  padding: 4px; box-shadow: 0 6px 24px rgba(0,0,0,.45); display: none;
}
.ctx-menu.on { display: block; }
.ctx-item { padding: 6px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; white-space: nowrap; }
.ctx-item:hover { background: var(--panel-2); }
.ctx-item.disabled { opacity: .4; cursor: default; }
.ctx-item.disabled:hover { background: none; }
.ctx-sep { height: 1px; background: var(--border); margin: 4px 2px; }
.btn {
  background: var(--panel-2); color: var(--text); border: 1px solid var(--border);
  border-radius: 4px; padding: 4px 10px; cursor: pointer; font-size: 12px; display: inline-flex; align-items: center; gap: 4px;
}
.btn:hover { background: #3a3a3d; }
.btn.primary { background: var(--accent-2); border-color: var(--accent-2); color: #fff; }
.btn.primary:hover { filter: brightness(1.1); }
.btn:disabled { opacity: .45; cursor: default; }

/* Main area */
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
/* Center via margin:auto on the child (.stage) instead of the flex container:
   a centered flex item that overflows the viewport makes the top/left unreachable by scrolling. */
.canvas-scroll { flex: 1 1 auto; overflow: auto; display: flex; padding: 20px; }
/* Slide content should not inherit the editor UI font settings (font size / family would change
   text metrics and line wrapping; align with examples/index.html defaults: 16px + sans-serif stack) */
.slide {
  font-size: 16px;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
}

/* Outer size = scaled size (ensures correct scroll range and centering); inner is 1:1 layout then scaled as a whole */
.stage { position: relative; background: #fff; box-shadow: 0 4px 20px rgba(0,0,0,.4); overflow: hidden; flex: 0 0 auto; margin: auto; }
.stage-inner { position: absolute; top: 0; left: 0; transform-origin: top left; }
.slide-host { position: absolute; top: 0; left: 0; user-select: none; -webkit-user-select: none; }
/* Inline-editing needs native text selection on the contenteditable body (parent disables it) */
.tb-body[contenteditable="true"] { user-select: text; -webkit-user-select: text; }
.slide-host .slide { margin: 0 !important; }
.overlay { position: absolute; top: 0; left: 0; pointer-events: none; }
.grid-overlay { position: absolute; inset: 0; pointer-events: none; display: none; }
.grid-overlay.on { display: block; background-image: linear-gradient(rgba(0,0,0,.08) 1px, transparent 1px), linear-gradient(90deg, rgba(0,0,0,.08) 1px, transparent 1px); background-size: 40px 40px; }

/* ===== Rendered slide elements =====
   Ported from the parser's editor renderer (examples/editor/styles.css) so elements
   built by render.ts get the same box model / text metrics as the reference editor. */
.slide-frame { position: absolute; top: 0; left: 0; }
.slide-frame.show-grid {
  background-image: linear-gradient(to right, rgba(0,0,0,.06) 1px, transparent 1px),
                    linear-gradient(to bottom, rgba(0,0,0,.06) 1px, transparent 1px);
}
.el { position: absolute; box-sizing: border-box; }
.el.locked { cursor: default; }
.el.hidden-el { display: none; }
.el-image img { width: 100%; height: 100%; display: block; }
.el-media video { pointer-events: none; }
.el.is-sel .el-media video { pointer-events: auto; }
.el-media .media-badge { display: flex; align-items: center; justify-content: center; }
.el.is-sel .el-media .media-badge { display: none; }
.el-text .tb-body {
  width: 100%; height: 100%; display: flex; flex-direction: column;
  justify-content: flex-start; outline: none; word-break: break-word;
  position: relative; z-index: 1;
  /* a:bodyPr lIns/tIns are inner padding and must sit inside the element box, otherwise a
     short text frame (e.g. 23.6px tall) loses its content area and the text drifts down. */
  box-sizing: border-box;
}
.el-text .tb-body > div { min-height: 1em; flex-shrink: 0; }
/* Hanging indent via negative first-line indent, so wrapped lines align with the bullet text */
.el-text .tb-body > div.bullet-para { padding-left: 1.3em; text-indent: -1.3em; }
.el-text .tb-body > div.num-para { padding-left: 1.8em; text-indent: -1.8em; }
.el-text .tb-body > div .bullet-mark {
  display: inline-block; min-width: 1.2em; text-align: left;
  white-space: pre; margin-right: 0.2em;
}
.el-text .tb-body > div .bullet-img img { display: inline-block; max-width: none; object-fit: contain; }
.el-table { display: grid; }
.el-table .cell {
  border: 1px solid #cbd5e1; padding: 4px 6px; overflow: hidden; display: flex;
  align-items: flex-start; word-break: break-word; white-space: pre-wrap;
}
.el-table .cell.cell-top { align-items: flex-start; }
.el-table .cell.cell-middle { align-items: center; }
.el-table .cell.cell-bottom { align-items: flex-end; }
.el-chart svg { display: block; width: 100%; height: 100%; }
.el-raw {
  border: 1px dashed #9aa0a6; background: repeating-linear-gradient(45deg,#f8f9fa,#f8f9fa 8px,#eceff1 8px,#eceff1 16px);
  display: flex; align-items: center; justify-content: center; color: var(--muted); font-size: 12px;
  text-align: center; padding: 6px;
}
/* Group edit state: dim other top-level elements and outline the active group */
.slide-host.group-editing .el.dimmed { opacity: 0.2; pointer-events: none; }
.slide-host.group-editing .el.group-edit-active { outline: 2px dashed var(--accent-2); outline-offset: -2px; }
.group-hint { position: absolute; left: 50%; top: 8px; transform: translateX(-50%); z-index: 30;
  background: var(--accent); color: #fff; padding: 4px 12px; border-radius: 14px; font-size: 12px;
  cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,.25); }
.el[data-editing="1"] .tb-body { outline: 2px solid var(--accent-2); }

/* Comment cards drawn on the canvas (matches the preview side's .pptx-comment styling) */
.slide-comment {
  position: absolute; width: 200px; box-sizing: border-box; z-index: 60;
  background: #fff7cc; border: 1px solid #f1d27a; border-radius: 8px;
  padding: 8px 10px; box-shadow: 0 2px 6px rgba(0,0,0,.15);
  font: 12px/1.4 sans-serif; color: #3b2f00; pointer-events: none;
}
.slide-comment .sc-author { font-weight: 700; margin-bottom: 2px; }
.slide-comment .sc-text { white-space: pre-wrap; }
.slide-comment .sc-date { margin-top: 4px; font-size: 10px; color: #8a7a3a; }

.sel-rect { position: absolute; border: 1px solid var(--accent); background: rgba(249,171,0,.08); pointer-events: auto; cursor: move; }
.sel-rect.selected { border-color: var(--accent-2); background: rgba(66,133,244,.1); }
.sel-rect.locked { border-style: dashed; cursor: default; }
.handle { position: absolute; width: 8px; height: 8px; background: #fff; border: 1px solid var(--accent-2); border-radius: 1px; }
.handle.nw { left: -4px; top: -4px; cursor: nwse-resize; }
.handle.n  { left: 50%; top: -4px; margin-left: -4px; cursor: ns-resize; }
.handle.ne { right: -4px; top: -4px; cursor: nesw-resize; }
.handle.e  { right: -4px; top: 50%; margin-top: -4px; cursor: ew-resize; }
.handle.se { right: -4px; bottom: -4px; cursor: nwse-resize; }
.handle.s  { left: 50%; bottom: -4px; margin-left: -4px; cursor: ns-resize; }
.handle.sw { left: -4px; bottom: -4px; cursor: nesw-resize; }
.handle.w  { left: -4px; top: 50%; margin-top: -4px; cursor: ew-resize; }
/* Rotation grip, positioned by JS above the top edge */
.handle.rot {
  left: 50%; margin-left: -5px; cursor: grab;
  width: 10px; height: 10px; border-radius: 50%;
  background: var(--accent-2); border-color: #fff;
}
/* Line endpoint handles (drag to resize/move an endpoint, snap to shapes as glue) */
.handle.endpoint {
  border-radius: 50%; background: var(--accent-2); border: 2px solid #fff;
  cursor: grab; pointer-events: auto; box-sizing: border-box;
}
.handle.endpoint:active { cursor: grabbing; }
/* Vertex points / bezier control points while in vertex-edit mode */
.handle.vertex {
  background: #fff; border: 1.5px solid var(--accent-2);
  cursor: move; pointer-events: auto; box-sizing: border-box;
}
.handle.vertex.sel { background: var(--accent-2); }
.handle.vctrl {
  border-radius: 50%; background: #fff; border: 1.5px solid var(--accent-2);
  cursor: move; pointer-events: auto; box-sizing: border-box;
}
/* Hit-area expansion: keeps the visible dot unchanged while enlarging the clickable region.
   Size is injected by JS as --hit (already divided by zoom → constant screen pixels). */
.handle::after {
  content: ''; position: absolute; left: 50%; top: 50%;
  width: var(--hit, 20px); height: var(--hit, 20px);
  transform: translate(-50%, -50%);
}
.lock-badge {
  position: absolute; top: -9px; left: -9px; font-size: 11px;
  background: var(--panel); border: 1px solid var(--border); border-radius: 4px; padding: 0 3px;
}
/* Snap guides while dragging */
.guide { position: absolute; background: #ea4335; pointer-events: none; }
.guide.h { height: 1px; left: 0; right: 0; }
.guide.v { width: 1px; top: 0; bottom: 0; }
/* Rubber-band selection box */
.marquee {
  position: absolute; border: 1px dashed var(--accent-2);
  background: rgba(66,133,244,.12); pointer-events: none;
}

/* Modal dialogs (tables, charts, shapes, media) */
.modal {
  position: fixed; inset: 0; background: rgba(0,0,0,.45);
  display: flex; align-items: center; justify-content: center; z-index: 60;
}
.dlg {
  background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
  box-shadow: 0 12px 32px rgba(0,0,0,.35); min-width: 400px; max-width: 92vw;
  max-height: 86vh; display: flex; flex-direction: column;
}
.dlg-head {
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
  padding: 10px 12px; border-bottom: 1px solid var(--border); font-weight: 600;
}
.dlg-x { cursor: pointer; opacity: .6; padding: 0 4px; }
.dlg-x:hover { opacity: 1; }
.dlg-body { padding: 12px; overflow: auto; display: flex; flex-direction: column; gap: 10px; }
.dlg-line { display: flex; align-items: center; gap: 8px; }
.dlg-line > span { width: 78px; flex: 0 0 auto; font-size: 12px; opacity: .8; }
.dlg-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.dlg-col { display: flex; flex-direction: column; gap: 6px; }
.dlg-block { border: 1px solid var(--border); border-radius: 6px; padding: 6px; margin-bottom: 6px; }
.dlg-split { display: flex; gap: 16px; }
.dlg-split > div { flex: 1 1 0; min-width: 0; }
.dlg-in, .dlg-cell {
  border: 1px solid var(--border); border-radius: 4px; padding: 3px 6px;
  font-size: 12px; background: transparent; color: inherit; min-width: 0;
}
.dlg-cell { width: 100%; box-sizing: border-box; }
.dlg-table { display: grid; gap: 4px; }
.dlg-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 6px; }
.dlg-check-row { display: flex; align-items: center; gap: 4px; font-size: 12px; }
.btn.primary { background: var(--accent-2); color: #fff; border-color: var(--accent-2); }
.dlg-shapes { display: grid; grid-template-columns: repeat(6, 1fr); gap: 6px; }
.shape-item {
  display: flex; flex-direction: column; align-items: center; gap: 2px;
  padding: 6px 2px; border: 1px solid var(--border); border-radius: 6px;
  background: transparent; color: inherit; cursor: pointer; font-size: 10px;
}
.shape-item:hover { border-color: var(--accent-2); background: rgba(66,133,244,.08); }
.dlg-group-title { font-size: 11px; color: var(--muted); font-weight: 600; margin-top: 2px; }

/* Right-side properties panel */
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

/* Status bar */
.statusbar { display: flex; align-items: center; gap: 8px; padding: 4px 10px; background: var(--panel); border-top: 1px solid var(--border); font-size: 12px; color: var(--muted); flex: 0 0 auto; }
.statusbar .spacer { flex: 1 1 auto; }
.chip { background: var(--panel-2); border: 1px solid var(--border); border-radius: 4px; padding: 2px 8px; cursor: pointer; }
.chip:hover { background: #3a3a3d; }

/* Presentation mode */
.present { position: fixed; inset: 0; background: #000; display: none; align-items: center; justify-content: center; z-index: 50; }
.present.on { display: flex; }
.present .stage { box-shadow: none; background: #000; }
.present .slide-host { position: relative; transform-origin: center center; }
.present .slide-host .slide { margin: 0 !important; }

/* Document properties panel (available in both preview and edit mode) */
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
