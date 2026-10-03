// PPTX editor webview main logic.
// Default "preview mode": renders original file bytes faithfully via pptxToHtml (same as examples/index.html).
// "Edit mode": switches to standard model round-trip rendering with a selectable/draggable/editable overlay.
import { CSS } from './style';
import { pptxToHtml } from '@fefeding/ppt-parser';
import { chartRenderer } from './vendor/chart-renderer';
import type { HostToWebview, WebviewToHost } from '../protocol';

declare function acquireVsCodeApi(): {
  postMessage(msg: WebviewToHost): void;
  getState(): any;
  setState(state: any): void;
};
const vscode = acquireVsCodeApi();

// ---------- Global state ----------
const state: {
  slideSize: { width: number; height: number };
  slidesHtml: string[]; // currently displayed HTML (preview or edit, filled after on-demand rendering)
  model: any;
  current: number;
  selected: { slide: number; element: number } | null;
  zoom: number;
  userZoom: number | null;
  editingText: boolean;
  grid: boolean;
  title: string;
  mode: 'preview' | 'edit';
  modelCanUndo: boolean;
  modelCanRedo: boolean;
} = {
  slideSize: { width: 1280, height: 720 },
  slidesHtml: [],
  model: null,
  current: 0,
  selected: null,
  zoom: 1,
  userZoom: null,
  editingText: false,
  grid: false,
  title: 'presentation.pptx',
  mode: 'preview',
  modelCanUndo: false,
  modelCanRedo: false
};

// Render sources (bytes from host) and render cache: on-demand rendering to avoid forcing preview to wait for model round-trip
let previewSrc: any = null; // original file bytes — faithful preview
let editSrc: any = null; // model round-tripped bytes — edit mode

/** A render result: per-slide HTML + chart data + document info (charts need a second pass by echarts). */
interface RenderResult {
  slides: string[];
  charts: any[];
  metadata?: any;
  customProps?: Record<string, string>;
}
const rendered: { preview?: RenderResult; edit?: RenderResult } = {};

// ---------- DOM references ----------
let root!: HTMLElement;
let slideListEl!: HTMLElement;
let stageEl!: HTMLElement;
let stageInnerEl!: HTMLElement;
let slideHostEl!: HTMLElement;
let overlayEl!: HTMLElement;
let gridEl!: HTMLElement;
let inspectorEl!: HTMLElement;
let ctxMenuEl!: HTMLElement;
let statusCountEl!: HTMLElement;
let zoomLabelEl!: HTMLElement;
let undoBtn!: HTMLButtonElement;
let redoBtn!: HTMLButtonElement;
let gridBtn!: HTMLElement;
let modeBtn!: HTMLButtonElement;
let presentEl!: HTMLElement;
let presentHostEl!: HTMLElement;
let docInfoEl!: HTMLElement;
let toastEl!: HTMLElement;

// ---------- Utility functions ----------
function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, any> = {},
  children: (Node | string)[] = []
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k === 'style') node.setAttribute('style', v);
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'html') node.innerHTML = v;
    else if (v != null) node.setAttribute(k, String(v));
  }
  for (const c of children) node.append(c);
  return node;
}

function num(v: any, def = 0): number {
  const n = typeof v === 'number' ? v : parseFloat(v);
  return Number.isFinite(n) ? n : def;
}
function clamp(v: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, v));
}
function post(msg: WebviewToHost) {
  vscode.postMessage(msg);
}
let toastTimer: any;
function toast(message: string, kind: 'info' | 'warn' | 'error' = 'info') {
  toastEl.textContent = message;
  toastEl.className = `toast show ${kind === 'info' ? '' : kind}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (toastEl.className = 'toast'), 2600);
}
function debounce(fn: (...a: any[]) => void, ms: number) {
  let t: any;
  return (...a: any[]) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...a), ms);
  };
}

// ---------- Layout ----------
function injectStyle() {
  const style = document.createElement('style');
  style.textContent = CSS;
  document.head.appendChild(style);
}

function buildLayout() {
  root = document.getElementById('app')!;
  root.className = 'app mode-preview';

  // Slide panel toolbar (replaces the top bar)
  const newBtn = h('button', { class: 'btn', title: 'New slide', onclick: () => sendOp({ kind: 'slideAdd', after: state.current }) }, ['+']);
  const presentBtn = h('button', { class: 'btn', title: 'Present (F5)', onclick: startPresent }, ['Present']);
  const docBtn = h('button', { class: 'btn', title: 'Document properties (metadata and custom properties)', onclick: toggleDocInfo }, ['Props']);
  modeBtn = h('button', { class: 'btn', title: 'Switch to edit mode', onclick: toggleMode }, ['Edit']);
  undoBtn = h('button', { class: 'btn edit-only', title: 'Undo (Ctrl+Z)', onclick: () => post({ type: 'undo' }) }, ['Undo']);
  redoBtn = h('button', { class: 'btn edit-only', title: 'Redo (Ctrl+Y)', onclick: () => post({ type: 'redo' }) }, ['Redo']);
  const toolbar = h('div', { class: 'sp-toolbar' }, [newBtn, presentBtn, docBtn, modeBtn, undoBtn, redoBtn]);

  // Slide list
  slideListEl = h('div', { class: 'sp-list' });
  const spHead = h('div', { class: 'sp-head' }, [h('span', {}, ['Slides'])]);
  const slidesPanel = h('div', { class: 'slides-panel' }, [spHead, toolbar, slideListEl]);

  // Canvas
  slideHostEl = h('div', { class: 'slide-host' });
  overlayEl = h('div', { class: 'overlay' });
  gridEl = h('div', { class: 'grid-overlay' });
  stageInnerEl = h('div', { class: 'stage-inner' }, [slideHostEl, gridEl, overlayEl]);
  stageEl = h('div', { class: 'stage' }, [stageInnerEl]);
  const scroll = h('div', { class: 'canvas-scroll', id: 'canvasScroll' }, [stageEl]);
  const canvasArea = h('div', { class: 'canvas-area' }, [scroll]);

  // Properties panel (edit mode only)
  inspectorEl = h('div', { class: 'inspector edit-only' });
  const main = h('div', { class: 'main' }, [slidesPanel, canvasArea, inspectorEl]);

  // Status bar
  statusCountEl = h('span', {}, ['0 slides']);
  zoomLabelEl = h('span', { class: 'chip', title: 'Click to reset to 100%', onclick: () => setZoom(1) }, ['100%']);
  const zoomOut = h('span', { class: 'chip', title: 'Zoom out', onclick: () => setZoom(state.userZoom ?? state.zoom, true, 0.9) }, ['−']);
  const zoomIn = h('span', { class: 'chip', title: 'Zoom in', onclick: () => setZoom(state.userZoom ?? state.zoom, true, 1.1) }, ['+']);
  const zoomFit = h('span', { class: 'chip', title: 'Fit to window', onclick: () => setZoom(null) }, ['Fit']);
  gridBtn = h('span', { class: 'chip edit-only', title: 'Grid', onclick: toggleGrid }, ['Grid']);
  const statusbar = h('div', { class: 'statusbar' }, [
    statusCountEl,
    h('span', { class: 'spacer' }),
    gridBtn,
    zoomOut, zoomFit, zoomLabelEl, zoomIn
  ]);

  // Presentation layer
  presentHostEl = h('div', { class: 'slide-host' });
  presentEl = h('div', { class: 'present', onclick: presentNext }, [presentHostEl]);

  docInfoEl = h('div', { class: 'doc-info' });

  toastEl = h('div', { class: 'toast' });

  // Right-click context menu (on canvas)
  ctxMenuEl = h('div', { class: 'ctx-menu' });
  canvasArea.addEventListener('contextmenu', (e) => { e.preventDefault(); showCtxMenu(e.clientX, e.clientY); });
  document.addEventListener('click', () => hideCtxMenu());
  window.addEventListener('scroll', () => hideCtxMenu(), true);

  root.append(main, statusbar, docInfoEl, presentEl, toastEl, ctxMenuEl);
}

// ---------- Mode switching ----------
function toggleMode() {
  const next: 'preview' | 'edit' = state.mode === 'preview' ? 'edit' : 'preview';
  state.mode = next;
  state.selected = null;
  post({ type: 'setMode', mode: next });
  applyMode();
}
function applyMode() {
  root.className = 'app mode-' + state.mode;
  modeBtn.textContent = state.mode === 'preview' ? 'Edit' : 'Preview';
  modeBtn.title = state.mode === 'preview' ? 'Switch to edit mode' : 'Switch to preview mode';

  const htmls = rendered[state.mode]?.slides;
  if (htmls) {
    state.slidesHtml = htmls;
    state.current = clamp(state.current, 0, state.slidesHtml.length - 1);
    renderAll();
    return;
  }
  // Cache not ready: render empty state first, then refresh when rendering completes (non-blocking)
  state.slidesHtml = [];
  renderAll();
  renderMode(state.mode);
}

/** Render the current mode; show toast on failure, never silently swallow errors. */
async function renderMode(mode: 'preview' | 'edit') {
  const src = mode === 'preview' ? previewSrc : editSrc;
  if (!src) return;
  try {
    // Skip silently when bytes for this mode aren't ready yet (e.g. preview mode: host doesn't do model round-trip)
    if (!toBytes(src).length) return;
    rendered[mode] = await renderFromBytes(src, mode === 'preview');
    if (state.mode === mode) {
      state.slidesHtml = rendered[mode]!.slides;
      state.current = clamp(state.current, 0, state.slidesHtml.length - 1);
      renderAll();
    }
  } catch (e: any) {
    console.error('[pptx-webview] Render failed:', e);
    toast('Render failed: ' + (e?.message || String(e)), 'error');
  }
}

// ---------- Zoom ----------
/** Fit to width (consistent with examples/index.html fitToWidth, capped at 2x). */
function getFitZoom(): number {
  const sc = document.getElementById('canvasScroll');
  const avail = (sc?.clientWidth || 800) - 40; // subtract left/right padding
  return clamp(avail / state.slideSize.width, 0.1, 2);
}
function setZoom(value: number | null, isUser = true, factor?: number) {
  if (value === null) {
    state.userZoom = null;
    state.zoom = getFitZoom();
  } else {
    let z = value;
    if (factor) z = (state.userZoom ?? state.zoom) * factor;
    z = clamp(z, 0.1, 4);
    state.userZoom = z;
    state.zoom = z;
  }
  applyStageTransform();
  zoomLabelEl.textContent = Math.round(state.zoom * 100) + '%';
}
function applyStageTransform() {
  stageInnerEl.style.width = state.slideSize.width + 'px';
  stageInnerEl.style.height = state.slideSize.height + 'px';
  stageInnerEl.style.transform = `scale(${state.zoom})`;
  // Outer placeholder = scaled size, to avoid clipping when zoomed in and blank scroll area when zoomed out
  stageEl.style.width = Math.round(state.slideSize.width * state.zoom) + 'px';
  stageEl.style.height = Math.round(state.slideSize.height * state.zoom) + 'px';
}

// ---------- Charts (echarts) ----------
/**
 * The same slide HTML appears in thumbnail / canvas / presentation layers, causing chart id duplication.
 * chart-renderer uses document.getElementById to find containers (only matches the first in the document).
 * So non-canvas copies get an id prefix to ensure the main canvas chart containers are unique.
 */
function prefixChartIds(html: string, prefix: string): string {
  return html.replace(/id=(["'])chart([^"']*)\1/g, `id=$1${prefix}chart$2$1`);
}

/** After containers are replaced, old echarts instances are detached from the document; dispose to avoid leaks. */
function disposeDetachedCharts() {
  const insts = (chartRenderer as any).chartInstances as Map<string, any> | undefined;
  if (!insts) return;
  for (const [id, inst] of Array.from(insts.entries())) {
    const dom = typeof inst?.getDom === 'function' ? inst.getDom() : null;
    if (!dom || !document.contains(dom)) {
      try {
        inst.dispose();
      } catch {
        /* noop */
      }
      insts.delete(id);
    }
  }
}

/** Paint charts that exist in the host using echarts (parser only emits empty placeholder divs). */
function paintCharts(host: HTMLElement, idPrefix = '') {
  const r = rendered[state.mode];
  if (!r || !r.charts.length) return;
  if (typeof (window as any).echarts === 'undefined') return; // skip silently if echarts not loaded
  const list = r.charts.filter((c) => !!host.querySelector('#' + idPrefix + c.chartId));
  if (!list.length) return;
  try {
    chartRenderer.renderCharts(
      idPrefix ? list.map((c) => ({ ...c, chartId: idPrefix + c.chartId })) : list
    );
  } catch (e: any) {
    console.warn('[pptx-webview] Chart rendering failed:', e);
  }
}

// ---------- Rendering ----------
function renderAll() {
  renderSlideList();
  renderCanvas();
  renderInspector();
  statusCountEl.textContent = `${state.slidesHtml.length} slides`;
  undoBtn.disabled = !state.modelCanUndo;
  redoBtn.disabled = !state.modelCanRedo;
  if (docInfoEl.classList.contains('on')) renderDocInfo();
}

function renderSlideList() {
  slideListEl.innerHTML = '';
  const thumbsW = 160;
  const scale = thumbsW / state.slideSize.width;
  state.slidesHtml.forEach((html, i) => {
    const inner = h('div', { class: 'inner', html: prefixChartIds(html, 'thumb-') });
    inner.style.width = state.slideSize.width + 'px';
    inner.style.height = state.slideSize.height + 'px';
    inner.style.transform = `scale(${scale})`;
    const thumb = h(
      'div',
      {
        class: 'thumb' + (i === state.current ? ' active' : ''),
        style: `width:${thumbsW}px;height:${thumbsW * (state.slideSize.height / state.slideSize.width)}px`,
        onclick: () => selectSlide(i)
      },
      [inner, h('span', { class: 'num' }, [String(i + 1)])]
    );
    slideListEl.append(thumb);
  });
}

function renderCanvas() {
  applyStageTransform();
  gridEl.className = 'grid-overlay' + (state.grid ? ' on' : '');
  disposeDetachedCharts();
  const html = state.slidesHtml[state.current] || '';
  slideHostEl.innerHTML = html;
  paintCharts(slideHostEl);
  if (state.mode === 'edit') buildOverlay();
  else overlayEl.innerHTML = '';
}

function buildOverlay() {
  overlayEl.innerHTML = '';
  const slide = state.model?.slides?.[state.current];
  if (!slide || !Array.isArray(slide.elements)) return;
  slide.elements.forEach((el: any, idx: number) => {
    if (el == null) return;
    const x = num(el.x), y = num(el.y), w = num(el.width), hgt = num(el.height);
    if (![x, y, w, hgt].every(Number.isFinite)) return;
    const rect = h('div', {
      class: 'sel-rect' + (state.selected?.slide === state.current && state.selected?.element === idx ? ' selected' : ''),
      style: `left:${x}px;top:${y}px;width:${w}px;height:${hgt}px`,
      'data-idx': String(idx)
    });
    const dirs = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
    for (const d of dirs) rect.append(h('div', { class: 'handle ' + d, 'data-dir': d }));
    rect.addEventListener('pointerdown', (e) => onElementPointerDown(e, idx));
    overlayEl.append(rect);
  });
}

function renderInspector() {
  // No properties panel in preview mode
  if (state.mode !== 'edit') {
    inspectorEl.innerHTML = '';
    return;
  }
  // Don't rebuild while editing text to avoid losing focus
  if (state.editingText && state.selected) return;
  inspectorEl.innerHTML = '';
  const slide = state.model?.slides?.[state.current];

  inspectorEl.append(h('h3', {}, ['Slide ' + (state.current + 1) + ' / ' + state.slidesHtml.length]));

  // Element list
  const list = h('ul', { class: 'el-list' });
  if (slide && Array.isArray(slide.elements)) {
    slide.elements.forEach((el: any, idx: number) => {
      const tag = elSnippet(el);
      const li = h(
        'li',
        {
          class: state.selected?.slide === state.current && state.selected?.element === idx ? 'active' : '',
          onclick: () => selectElement(idx)
        },
        [h('span', {}, [tag.name]), h('span', { class: 'tag' }, [tag.type])]
      );
      list.append(li);
    });
  }
  inspectorEl.append(h('h4', {}, ['Elements (click to select)']), list);

  if (!state.selected) {
    renderSlideInspector(slide);
  } else {
    const el = slide?.elements?.[state.selected.element];
    if (!el) {
      state.selected = null;
      renderSlideInspector(slide);
    } else {
      renderElementInspector(el);
    }
  }
}

function elSnippet(el: any): { name: string; type: string } {
  switch (el.type) {
    case 'text':
      return { name: (el.text || '').toString().split('\n')[0].slice(0, 24) || 'Text', type: 'Text' };
    case 'shape':
      return { name: (el.shapeType || 'shape').toString(), type: 'Shape' };
    case 'image':
      return { name: 'Image', type: 'Image' };
    case 'chart':
      return { name: (el.chartType || 'chart').toString(), type: 'Chart' };
    case 'table':
      return { name: 'Table', type: 'Table' };
    case 'group':
      return { name: 'Group', type: 'Group' };
    case 'diagram':
      return { name: 'Diagram', type: 'Diagram' };
    case 'video':
      return { name: 'Video', type: 'Video' };
    case 'audio':
      return { name: 'Audio', type: 'Audio' };
    default:
      return { name: (el.type || 'element').toString(), type: el.type || 'Element' };
  }
}

// ---- Slide-level properties ----
function renderSlideInspector(slide: any) {
  if (!slide) return;
  inspectorEl.append(h('h4', {}, ['Slide Properties']));

  const bg = colorField('Background', slide.background && typeof slide.background === 'string' ? slide.background : '#ffffff', (v) =>
    sendOp({ kind: 'slideUpdate', index: state.current, patch: { background: v } })
  );
  inspectorEl.append(bg);

  const hidden = h('input', { type: 'checkbox', ...(slide.hidden ? { checked: 'checked' } : {}) });
  hidden.addEventListener('change', () => sendOp({ kind: 'slideUpdate', index: state.current, patch: { hidden: hidden.checked } }));
  inspectorEl.append(h('div', { class: 'field' }, [h('label', {}, ['Hidden']), hidden]));

  const notes = h('textarea', { placeholder: 'Speaker notes...' }, [slide.notes || '']);
  notes.addEventListener('input', debounce(() => sendOp({ kind: 'slideUpdate', index: state.current, patch: { notes: notes.value } }), 400));
  inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, ['Notes']), notes]));

  // Slide actions
  inspectorEl.append(h('h4', {}, ['Slide Actions']));
  const addText = h('button', { class: 'btn', onclick: () => addTextElement() }, ['+ Text Box']);
  const addShape = h('button', { class: 'btn', onclick: () => addShapeElement('rect') }, ['+ Rectangle']);
  const addImg = h('button', { class: 'btn', onclick: pickImage }, ['+ Image']);
  inspectorEl.append(h('div', { class: 'row' }, [addText, addShape, addImg]));

  const dup = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'slideDuplicate', index: state.current }) }, ['Duplicate']);
  const del = h('button', { class: 'btn', onclick: () => { if (confirm('Delete this slide?')) sendOp({ kind: 'slideDelete', index: state.current }); } }, ['Delete']);
  const up = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'slideMove', from: state.current, to: Math.max(0, state.current - 1) }) }, ['Move Up']);
  const down = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'slideMove', from: state.current, to: Math.min(state.slidesHtml.length - 1, state.current + 1) }) }, ['Move Down']);
  inspectorEl.append(h('div', { class: 'row' }, [dup, del]));
  inspectorEl.append(h('div', { class: 'row' }, [up, down]));
}

// ---- Element-level properties ----
function renderElementInspector(el: any) {
  inspectorEl.append(h('h4', {}, ['Element: ' + (el.type || 'unknown')]));

  // Geometry
  const geom = h('div', {});
  geom.append(numField('X', el.x, (v) => updateEl({ x: v })));
  geom.append(numField('Y', el.y, (v) => updateEl({ y: v })));
  geom.append(numField('Width', el.width, (v) => updateEl({ width: Math.max(1, v) })));
  geom.append(numField('Height', el.height, (v) => updateEl({ height: Math.max(1, v) })));
  geom.append(numField('Rotation', el.rotation, (v) => updateEl({ rotation: v })));
  inspectorEl.append(geom);

  // Type-specific
  if (el.type === 'text') {
    const ta = h('textarea', { placeholder: 'Text content...' }, [el.text || '']);
    ta.addEventListener('focus', () => (state.editingText = true));
    ta.addEventListener('blur', () => { state.editingText = false; });
    ta.addEventListener('input', debounce(() => { updateElLocal({ text: ta.value }); sendOp({ kind: 'elementUpdate', slide: state.current, element: state.selected!.element, patch: { text: ta.value } }); }, 400));
    inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, ['Text']), ta]));

    inspectorEl.append(colorField('Color', el.color || '#1e1e1e', (v) => updateEl({ color: v })));
    inspectorEl.append(numField('Font Size', el.fontSize, (v) => updateEl({ fontSize: v })));
    inspectorEl.append(selectField('Align', ['left', 'center', 'right', 'justify'], el.align || 'left', (v) => updateEl({ align: v })));
    inspectorEl.append(selectField('Vertical', ['top', 'middle', 'bottom'], el.valign || 'top', (v) => updateEl({ valign: v })));
    inspectorEl.append(checkField('Bold', !!el.bold, (v) => updateEl({ bold: v })));
    inspectorEl.append(checkField('Italic', !!el.italic, (v) => updateEl({ italic: v })));
    inspectorEl.append(checkField('Underline', !!el.underline, (v) => updateEl({ underline: v })));
  } else if (el.type === 'shape') {
    inspectorEl.append(textField('Shape', el.shapeType || 'rect', (v) => updateEl({ shapeType: v })));
    const fillVal = typeof el.fill === 'string' ? el.fill : (el.fill && el.fill.color) || '#4285f4';
    inspectorEl.append(colorField('Fill', fillVal, (v) => updateEl({ fill: v })));
    const lineColor = el.line && el.line !== 'none' ? (el.line.color || '#000') : '#000';
    const lineWidth = el.line && el.line !== 'none' ? (el.line.width || 1) : 1;
    inspectorEl.append(colorField('Border Color', lineColor, (v) => updateEl({ line: { color: v, width: lineWidth } })));
    inspectorEl.append(numField('Border Width', lineWidth, (v) => updateEl({ line: v > 0 ? { color: lineColor, width: v } : 'none' })));
  } else if (el.type === 'image') {
    const replace = h('button', { class: 'btn', onclick: pickImage }, ['Replace Image']);
    inspectorEl.append(h('div', { class: 'row' }, [replace]));
    inspectorEl.append(h('div', { class: 'empty' }, ['Image content comes from preview. Select a local file to replace.']));
  } else {
    inspectorEl.append(h('div', { class: 'empty' }, ['This type only supports move / resize / delete']));
  }

  // Element actions
  inspectorEl.append(h('h4', {}, ['Element Actions']));
  const del = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'elementDelete', slide: state.current, element: state.selected!.element }) }, ['Delete']);
  const fwd = h('button', { class: 'btn', onclick: () => reorderEl(1) }, ['Bring Forward']);
  const bwd = h('button', { class: 'btn', onclick: () => reorderEl(-1) }, ['Send Backward']);
  inspectorEl.append(h('div', { class: 'row' }, [del]));
  inspectorEl.append(h('div', { class: 'row' }, [fwd, bwd]));
}

// ---------- Field builders ----------
function colorField(label: string, value: string, onInput: (v: string) => void) {
  const input = h('input', { type: 'color', value: toHex(value) });
  input.addEventListener('input', () => onInput(input.value));
  const clear = h('button', { class: 'btn', style: 'padding:2px 6px', title: 'Set to none', onclick: () => onInput('none') }, ['None']);
  return h('div', { class: 'field' }, [h('label', {}, [label]), input, clear]);
}
function textField(label: string, value: string, onInput: (v: string) => void) {
  const input = h('input', { type: 'text', value: value ?? '' });
  input.addEventListener('change', () => onInput(input.value));
  return h('div', { class: 'field' }, [h('label', {}, [label]), input]);
}
function numField(label: string, value: any, onInput: (v: number) => void) {
  const input = h('input', { type: 'number', value: num(value) });
  input.addEventListener('change', () => onInput(num(input.value)));
  return h('div', { class: 'field' }, [h('label', {}, [label]), input]);
}
function selectField(label: string, options: string[], value: string, onInput: (v: string) => void) {
  const sel = h('select', {});
  for (const o of options) sel.append(h('option', { value: o, ...(o === value ? { selected: 'selected' } : {}) }, [o]));
  sel.addEventListener('change', () => onInput(sel.value));
  return h('div', { class: 'field' }, [h('label', {}, [label]), sel]);
}
function checkField(label: string, value: boolean, onInput: (v: boolean) => void) {
  const input = h('input', { type: 'checkbox', ...(value ? { checked: 'checked' } : {}) });
  input.addEventListener('change', () => onInput(input.checked));
  return h('div', { class: 'field' }, [h('label', {}, [label]), input]);
}
function toHex(v: any): string {
  if (typeof v === 'string' && /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(v)) return v;
  return '#ffffff';
}

// ---------- Selection / Operations ----------
function selectSlide(i: number) {
  state.current = clamp(i, 0, state.slidesHtml.length - 1);
  state.selected = null;
  renderSlideList();
  renderCanvas();
  renderInspector();
}
function selectElement(idx: number) {
  if (state.mode !== 'edit') return;
  state.selected = { slide: state.current, element: idx };
  renderSlideList();
  renderCanvas();
  renderInspector();
}
function sendOp(op: any) {
  post({ type: 'op', op });
}
function updateEl(patch: any) {
  if (!state.selected) return;
  updateElLocal(patch);
  sendOp({ kind: 'elementUpdate', slide: state.selected.slide, element: state.selected.element, patch });
}
function updateElLocal(patch: any) {
  if (!state.selected) return;
  const el = state.model?.slides?.[state.selected.slide]?.elements?.[state.selected.element];
  if (el) Object.assign(el, patch);
}

function reorderEl(dir: number) {
  if (!state.selected) return;
  const from = state.selected.element;
  const to = from + dir;
  const len = state.model.slides[state.selected.slide].elements.length;
  if (to < 0 || to >= len) return;
  sendOp({ kind: 'elementReorder', slide: state.selected.slide, from, to });
  state.selected.element = to;
}

function addTextElement() {
  sendOp({
    kind: 'elementAdd',
    slide: state.current,
    element: { type: 'text', x: 120, y: 120, width: 480, height: 90, text: 'Double-click to edit text', fontSize: 24, color: '#1e1e1e', align: 'left', valign: 'top' }
  });
}
function addShapeElement(shapeType: string) {
  sendOp({
    kind: 'elementAdd',
    slide: state.current,
    element: { type: 'shape', shapeType, x: 200, y: 200, width: 200, height: 120, fill: '#4285f4', line: { color: '#000', width: 1 } }
  });
}
function pickImage() {
  const input = h('input', { type: 'file', accept: 'image/*', style: 'display:none' });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result as string;
      const img = new Image();
      img.onload = () => {
        const w = 360;
        const hgt = Math.round((img.height / img.width) * w) || 240;
        sendOp({ kind: 'elementAdd', slide: state.current, element: { type: 'image', data: dataUrl, x: 160, y: 160, width: w, height: hgt } });
      };
      img.src = dataUrl;
    };
    reader.readAsDataURL(file);
  });
  document.body.append(input);
  input.click();
}

// ---------- Drag / Resize ----------
function onElementPointerDown(e: PointerEvent, idx: number) {
  e.preventDefault();
  e.stopPropagation();
  state.selected = { slide: state.current, element: idx };
  renderInspector();
  buildOverlaySelected();

  const el = state.model.slides[state.current].elements[idx];
  const dir = (e.target as HTMLElement).getAttribute('data-dir');
  const startX = e.clientX, startY = e.clientY;
  const orig = { x: num(el.x), y: num(el.y), w: num(el.width), h: num(el.height) };
  const rectEl = (e.currentTarget as HTMLElement);
  rectEl.setPointerCapture(e.pointerId);

  const onMove = (ev: PointerEvent) => {
    const dx = (ev.clientX - startX) / state.zoom;
    const dy = (ev.clientY - startY) / state.zoom;
    let nx = orig.x, ny = orig.y, nw = orig.w, nh = orig.h;
    if (!dir) {
      nx = orig.x + dx; ny = orig.y + dy;
    } else {
      if (dir.includes('e')) nw = Math.max(10, orig.w + dx);
      if (dir.includes('s')) nh = Math.max(10, orig.h + dy);
      if (dir.includes('w')) { nx = orig.x + dx; nw = Math.max(10, orig.w - dx); }
      if (dir.includes('n')) { ny = orig.y + dy; nh = Math.max(10, orig.h - dy); }
    }
    Object.assign(el, { x: nx, y: ny, width: nw, height: nh });
    rectEl.style.left = nx + 'px';
    rectEl.style.top = ny + 'px';
    rectEl.style.width = nw + 'px';
    rectEl.style.height = nh + 'px';
  };
  const onUp = (ev: PointerEvent) => {
    rectEl.releasePointerCapture(ev.pointerId);
    rectEl.removeEventListener('pointermove', onMove);
    rectEl.removeEventListener('pointerup', onUp);
    sendOp({ kind: 'elementUpdate', slide: state.current, element: idx, patch: { x: num(el.x), y: num(el.y), width: num(el.width), height: num(el.height) } });
    renderInspector();
  };
  rectEl.addEventListener('pointermove', onMove);
  rectEl.addEventListener('pointerup', onUp);
}
function buildOverlaySelected() {
  overlayEl.querySelectorAll('.sel-rect').forEach((r) => {
    const idx = Number((r as HTMLElement).getAttribute('data-idx'));
    (r as HTMLElement).className = 'sel-rect' + (state.selected?.element === idx ? ' selected' : '');
  });
}

// ---------- Grid / Presentation ----------
function toggleGrid() {
  state.grid = !state.grid;
  gridEl.className = 'grid-overlay' + (state.grid ? ' on' : '');
  gridBtn.style.background = state.grid ? 'var(--accent)' : '';
}
function toggleDocInfo() {
  const on = docInfoEl.classList.toggle('on');
  if (on) renderDocInfo();
}

const META_LABELS: Record<string, string> = {
  title: 'Title',
  subject: 'Subject',
  author: 'Author',
  keywords: 'Keywords',
  description: 'Description',
  lastModifiedBy: 'Last Modified By',
  created: 'Created',
  modified: 'Modified',
  category: 'Category',
  status: 'Status',
  contentType: 'Content Type',
  language: 'Language'
};

function kvTable(obj: any, labels?: Record<string, string>) {
  const keys = Object.keys(obj || {}).filter((k) => obj[k] != null && obj[k] !== '');
  if (!keys.length) return h('div', { class: 'empty' }, ['(none)']);
  const table = h('table', { class: 'kv' });
  for (const k of keys) {
    table.append(
      h('tr', {}, [
        h('td', { class: 'k' }, [labels?.[k] || k]),
        h('td', { class: 'v' }, [String(obj[k])])
      ])
    );
  }
  return table;
}

/** Render document properties: metadata (core.xml) + custom properties (custom.xml) */
function renderDocInfo() {
  const r = rendered[state.mode];
  docInfoEl.innerHTML = '';
  docInfoEl.append(
    h('div', { class: 'doc-info-head' }, [
      h('span', {}, ['Document Properties']),
      h('button', { class: 'btn', style: 'padding:2px 8px', onclick: () => docInfoEl.classList.remove('on') }, ['Close'])
    ])
  );
  docInfoEl.append(h('h4', {}, ['Metadata']), kvTable(r?.metadata, META_LABELS));
  docInfoEl.append(h('h4', {}, ['Custom Properties']), kvTable(r?.customProps));
}

function startPresent() {
  presentEl.className = 'present on';
  renderPresent();
}
function renderPresent() {
  disposeDetachedCharts();
  presentHostEl.style.width = state.slideSize.width + 'px';
  presentHostEl.style.height = state.slideSize.height + 'px';
  presentHostEl.innerHTML = prefixChartIds(state.slidesHtml[state.current] || '', 'pres-');
  paintCharts(presentHostEl, 'pres-');
  const z = Math.min(window.innerWidth / state.slideSize.width, window.innerHeight / state.slideSize.height) * 0.96;
  presentHostEl.style.transform = `scale(${z})`;
}
function presentNext() {
  if (state.current < state.slidesHtml.length - 1) {
    state.current++;
    renderPresent();
  } else {
    endPresent();
  }
}
function endPresent() {
  presentEl.className = 'present';
}
function presentPrev() {
  if (state.current > 0) { state.current--; renderPresent(); }
}

// ---------- Context menu ----------
type CtxItem = { label?: string; action?: () => void; disabled?: boolean; sep?: boolean };
function showCtxMenu(x: number, y: number) {
  ctxMenuEl.innerHTML = '';
  const items: CtxItem[] = [
    { label: 'Present (F5)', action: startPresent },
    { label: state.mode === 'preview' ? 'Edit' : 'Preview', action: toggleMode },
    { label: 'Document Properties', action: toggleDocInfo },
    { label: 'Save (Ctrl+S)', action: () => post({ type: 'save' }) },
    { label: 'Export Copy...', action: () => post({ type: 'saveAs' }) },
    { sep: true },
    { label: 'Undo (Ctrl+Z)', action: () => post({ type: 'undo' }), disabled: !state.modelCanUndo },
    { label: 'Redo (Ctrl+Y)', action: () => post({ type: 'redo' }), disabled: !state.modelCanRedo }
  ];
  for (const it of items) {
    if (it.sep) { ctxMenuEl.append(h('div', { class: 'ctx-sep' })); continue; }
    const item = h('div', { class: 'ctx-item' + (it.disabled ? ' disabled' : '') }, [it.label]);
    if (!it.disabled && it.action) {
      item.addEventListener('click', (e) => {
        e.stopPropagation();
        hideCtxMenu();
        it.action!();
      });
    }
    ctxMenuEl.append(item);
  }
  ctxMenuEl.classList.add('on');
  const mw = ctxMenuEl.offsetWidth || 180;
  const mh = ctxMenuEl.offsetHeight || 200;
  const left = Math.min(x, window.innerWidth - mw - 8);
  const top = Math.min(y, window.innerHeight - mh - 8);
  ctxMenuEl.style.left = Math.max(4, left) + 'px';
  ctxMenuEl.style.top = Math.max(4, top) + 'px';
}
function hideCtxMenu() {
  ctxMenuEl.classList.remove('on');
}

// ---------- Message handling ----------
/** Normalize host-provided bytes to Uint8Array (handles TypedArray / ArrayBuffer / array / base64 / plain object fallback). */
function toBytes(v: any): Uint8Array {
  if (v instanceof Uint8Array) return v;
  if (v instanceof ArrayBuffer) return new Uint8Array(v);
  if (ArrayBuffer.isView(v)) {
    const t = v as ArrayBufferView;
    return new Uint8Array(t.buffer, t.byteOffset, t.byteLength);
  }
  if (Array.isArray(v)) return Uint8Array.from(v as number[]);
  if (typeof v === 'string') {
    const bin = atob(v);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  if (v && typeof v === 'object') {
    // postMessage fallback as plain object: {"0":80,"1":75,...}
    const keys = Object.keys(v).filter((k) => /^\d+$/.test(k));
    if (keys.length) {
      const out = new Uint8Array(keys.length);
      for (const k of keys) out[Number(k)] = Number(v[k]) & 0xff;
      return out;
    }
  }
  throw new Error('Unrecognized PPTX byte data');
}

/** Inject parser-produced global styles. Missing it causes colors/line-height/SVG filters to be lost (rendering corruption). */
let injectedGlobalCss = '';
function ensureGlobalStyles(css: string) {
  if (!css || css === injectedGlobalCss) return;
  let el = document.getElementById('pptx-global-css') as HTMLStyleElement | null;
  if (!el) {
    el = document.createElement('style');
    el.id = 'pptx-global-css';
    document.head.appendChild(el);
  }
  el.textContent = css;
  injectedGlobalCss = css;
}

/**
 * Render bytes into per-slide HTML using pptxToHtml.
 * Consistent with examples/index.html: injects styles.global; preview mode skips hidden slides (p:sld show="0").
 */
async function renderFromBytes(bytes: any, skipHidden = false): Promise<RenderResult> {
  const data = toBytes(bytes);
  if (!data.length) throw new Error('PPTX bytes are empty, cannot parse');
  const res = await pptxToHtml(data, { mediaProcess: true, themeProcess: true });
  if (!res) throw new Error('pptxToHtml returned no result');
  ensureGlobalStyles(res.styles?.global || '');
  if (res.slideSize && res.slideSize.width) {
    state.slideSize = { width: res.slideSize.width, height: res.slideSize.height };
  }
  const slides = (res.slides || []) as any[];
  return {
    slides: (skipHidden ? slides.filter((s) => !s.hidden) : slides).map((s) => s.html),
    charts: (res.charts || []) as any[],
    metadata: res.metadata,
    customProps: res.customProps
  };
}

async function applyInit(msg: Extract<HostToWebview, { type: 'init' }>) {
  state.model = msg.model;
  state.title = msg.title;
  state.current = 0;
  state.selected = null;
  state.mode = msg.mode || 'preview';
  state.modelCanUndo = msg.canUndo;
  state.modelCanRedo = msg.canRedo;

  // Store render sources; only render the current mode (preview uses original bytes for faithful rendering, edit uses model round-trip result)
  previewSrc = msg.originalBytes;
  editSrc = msg.bytes;
  delete rendered.preview;
  delete rendered.edit;
  await renderMode(state.mode);

  applyMode();
  setZoom(null);

}
async function applyUpdate(msg: Extract<HostToWebview, { type: 'update' }>) {
  state.model = msg.model;
  state.mode = msg.mode || state.mode;
  state.modelCanUndo = msg.canUndo;
  state.modelCanRedo = msg.canRedo;
  previewSrc = msg.originalBytes;
  editSrc = msg.bytes;
  delete rendered.preview;
  delete rendered.edit;
  await renderMode(state.mode);

  state.slidesHtml = rendered[state.mode]?.slides || [];
  state.current = clamp(state.current, 0, state.slidesHtml.length - 1);
  if (state.selected && state.selected.slide === state.current) {
    const len = state.model?.slides?.[state.current]?.elements?.length ?? 0;
    if (state.selected.element >= len) state.selected = null;
  }
  applyMode();
}

// ---------- Keyboard ----------
function onKey(e: KeyboardEvent) {
  const tag = (e.target as HTMLElement)?.tagName;
  const typing = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';
  if (presentEl.classList.contains('on')) {
    if (e.key === 'Escape') endPresent();
    else if (e.key === 'ArrowRight' || e.key === ' ') presentNext();
    else if (e.key === 'ArrowLeft') presentPrev();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); post({ type: 'save' }); return; }
  if (state.mode !== 'edit') return; // disable edit shortcuts in preview mode
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); post({ type: 'undo' }); return; }
  if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) { e.preventDefault(); post({ type: 'redo' }); return; }
  if ((e.key === 'F5')) { e.preventDefault(); startPresent(); return; }
  if (typing) return;
  if ((e.key === 'Delete' || e.key === 'Backspace') && state.selected) {
    e.preventDefault();
    sendOp({ kind: 'elementDelete', slide: state.selected.slide, element: state.selected.element });
    state.selected = null;
  } else if (e.key === 'Escape') {
    state.selected = null; renderInspector(); buildOverlaySelected();
  } else if ((e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') && state.selected) {
    e.preventDefault();
    const step = e.shiftKey ? 10 : 1;
    const el = state.model.slides[state.selected.slide].elements[state.selected.element];
    const patch: any = {};
    if (e.key === 'ArrowUp') patch.y = num(el.y) - step;
    if (e.key === 'ArrowDown') patch.y = num(el.y) + step;
    if (e.key === 'ArrowLeft') patch.x = num(el.x) - step;
    if (e.key === 'ArrowRight') patch.x = num(el.x) + step;
    updateEl(patch);
    buildOverlay();
  }
}

// ---------- Startup ----------
function main() {
  injectStyle();
  buildLayout();
  window.addEventListener('keydown', onKey);
  window.addEventListener('resize', () => { if (!state.userZoom) setZoom(null); });
  window.addEventListener('message', (ev) => {
    const msg = ev.data as HostToWebview;
    if (!msg) return;
    (async () => {
      try {
        switch (msg.type) {
          case 'init': await applyInit(msg as any); break;
          case 'update': await applyUpdate(msg as any); break;
          case 'saved': break;
          case 'info': toast(msg.message, msg.kind || 'info'); break;
        }
      } catch (e: any) {
        // Never silently swallow errors: otherwise the UI would stall at "0 slides" with no feedback
        console.error('[pptx-webview] Message handling failed:', e);
        toast('Render failed: ' + (e?.message || String(e)), 'error');
      }
    })();
  });
  post({ type: 'ready' });
}

main();
