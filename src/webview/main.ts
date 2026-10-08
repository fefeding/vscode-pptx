// PPTX editor webview main logic.
//
// Preview mode: renders the original file bytes with pptxToHtml (identical to the parser's
// preview page) — always the most faithful view, also used for presenting.
// Edit mode: parses the bytes into the parser's editor document (pptxToStandard -> docFromPptx)
// and renders straight from that model (no PPTX byte round-trip, so nothing is lost in
// serialization). All edits go through the parser's editor core (createStore / createActions),
// which owns the model, selection and undo/redo history.
//
// The extension host only stores PPTX bytes: after each edit the webview serializes the
// document and syncs the result, so the host can save without understanding the model.
import { CSS } from './style';
import { h, clamp } from './util';
import {
  pptxToHtml,
  pptxToStandard,
  jsonToPptx,
  docFromPptx,
  docToPptx,
  createStore,
  createActions,
  elementRect,
  effectMargin,
  FONT_LIST,
  THEMES,
  LAYOUTS,
  SLIDE_SIZES,
  TRANSITIONS,
  ANIM_CLASSES,
  ANIM_TYPES
} from '@fefeding/ppt-parser';
import {
  renderSlideInto,
  renderThumbInto,
  disposeDetachedCharts,
  disposeAllCharts
} from './render';
import {
  hasOpenModal,
  closeModal,
  openTableSizeDialog,
  openTableDialog,
  openChartDialog,
  openShapePicker,
  openMediaDialog
} from './dialogs';
// Same ECharts renderer the preview side uses (option building incl. 3D matches examples/index.html)
import { chartRenderer } from '@fefeding/ppt-parser/chart-renderer';
import type { EditorMode, HostToWebview, WebviewToHost } from '../protocol';

declare function acquireVsCodeApi(): {
  postMessage(msg: WebviewToHost): void;
  getState(): any;
  setState(state: any): void;
};
const vscode = acquireVsCodeApi();

const PARSE_OPTS = { mediaProcess: true, themeProcess: true } as const;

// ---------- Document state ----------
// One editor store per webview (the parser core keeps the document, selection and history).
const store: any = createStore();
const actions: any = createActions(store);

const state: {
  mode: EditorMode;
  title: string;
  /** Original file bytes — preview renders these. */
  original: Uint8Array | null;
  /** Latest bytes: original until the webview reports an edit. */
  latest: Uint8Array | null;
  /** Preview render result (pptxToHtml output). */
  preview: { slides: string[]; charts: any[]; metadata?: any; customProps?: Record<string, string> } | null;
  present: { slides: string[]; charts: any[]; version: number } | null;
  /** Bumped whenever the serialized document changes, to invalidate cached renders. */
  version: number;
  slideSize: { width: number; height: number };
  current: number;
  zoom: number;
  userZoom: number | null;
  grid: boolean;
  editingText: boolean;
  syncing: boolean;
} = {
  mode: 'preview',
  title: 'presentation.pptx',
  original: null,
  latest: null,
  preview: null,
  present: null,
  version: 0,
  slideSize: { width: 1280, height: 720 },
  current: 0,
  zoom: 1,
  userZoom: null,
  grid: false,
  editingText: false,
  syncing: false
};

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
let modeBtn!: HTMLElement;
let presentEl!: HTMLElement;
let presentHostEl!: HTMLElement;
let docInfoEl!: HTMLElement;
let toastEl!: HTMLElement;

// ---------- Utilities ----------
function num(v: any, def = 0): number {
  const n = typeof v === 'number' ? v : parseFloat(v);
  return Number.isFinite(n) ? n : def;
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
function toBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function toBase64(bytes: Uint8Array): string {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)) as any);
  }
  return btoa(bin);
}

// ---------- Layout ----------
function buildLayout() {
  root = document.getElementById('app')!;
  root.className = 'app mode-preview';

  // Slide panel toolbar
  const newBtn = h('button', { class: 'btn', title: 'New slide', onclick: () => runAction(() => actions.addSlide()) }, ['+']);
  const presentBtn = h('button', { class: 'btn', title: 'Present (F5)', onclick: startPresent }, ['Present']);
  const docBtn = h('button', { class: 'btn', title: 'Document properties (metadata and custom properties)', onclick: toggleDocInfo }, ['Props']);
  modeBtn = h('button', { class: 'btn', title: 'Switch to edit mode', onclick: toggleMode }, ['Edit']);
  undoBtn = h('button', { class: 'btn edit-only', title: 'Undo (Ctrl+Z)', onclick: () => runAction(() => store.undo()) }, ['Undo']);
  redoBtn = h('button', { class: 'btn edit-only', title: 'Redo (Ctrl+Y)', onclick: () => runAction(() => store.redo()) }, ['Redo']);
  const saveBtn = h('button', { class: 'btn edit-only', title: 'Save', onclick: flushAndSave }, ['Save']);
  // Quick insert row (mirrors the editor's insert toolbar)
  const insText = h('button', { class: 'btn edit-only', title: 'Insert text box', onclick: addTextElement }, ['T']);
  const insShape = h('button', { class: 'btn edit-only', title: 'Insert shape', onclick: () => openShapePicker({ addShape: (t: string) => addShapeElement(t) }) }, ['▢']);
  const insImg = h('button', { class: 'btn edit-only', title: 'Insert image', onclick: pickImage }, ['▣']);
  const insTable = h('button', { class: 'btn edit-only', title: 'Insert table', onclick: insertTable }, ['⊞']);
  const insChart = h('button', { class: 'btn edit-only', title: 'Insert chart', onclick: insertChart }, ['◫']);
  const insMedia = h('button', { class: 'btn edit-only', title: 'Insert audio or video', onclick: () => openMediaDialog({ addMedia }) }, ['♪']);
  const toolbar = h('div', { class: 'sp-toolbar' }, [
    newBtn, presentBtn, docBtn, modeBtn, saveBtn, undoBtn, redoBtn,
    insText, insShape, insImg, insTable, insChart, insMedia
  ]);

  // Slide list
  slideListEl = h('div', { class: 'sp-list' });
  const spHead = h('div', { class: 'sp-head' }, [h('span', {}, ['Slides'])]);
  const slidesPanel = h('div', { class: 'slides-panel' }, [spHead, toolbar, slideListEl]);

  // Canvas: frame + overlay share one scaled coordinate space
  slideHostEl = h('div', { class: 'slide-host slide-frame' });
  overlayEl = h('div', { class: 'overlay' });
  gridEl = h('div', { class: 'grid-overlay' });
  stageInnerEl = h('div', { class: 'stage-inner' }, [slideHostEl, gridEl, overlayEl]);
  stageEl = h('div', { class: 'stage' }, [stageInnerEl]);
  const scroll = h('div', { class: 'canvas-scroll', id: 'canvasScroll' }, [stageEl]);
  // Ctrl/Cmd + wheel zooms the canvas (plain wheel still scrolls)
  scroll.addEventListener(
    'wheel',
    (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      setZoom(state.userZoom ?? state.zoom, true, e.deltaY < 0 ? 1.1 : 0.9);
    },
    { passive: false }
  );
  // Canvas interaction: hit-test to select, drag handles live on the overlay boxes.
  slideHostEl.addEventListener('pointerdown', (e) => {
    if (state.mode !== 'edit') return;
    const node = (e.target as HTMLElement).closest('.el[data-id]') as HTMLElement | null;
    if (node) {
      const id = node.dataset.id!;
      if (store.editingId === id) return; // let contenteditable place the caret itself
      e.preventDefault();
      if (e.shiftKey) store.toggleSel(id);
      else if (!store.sel.includes(id)) store.setSel([id]);
      renderSelection();
      renderInspector();
      if (!e.shiftKey) startMove(e);
      return;
    }
    if (store.editingId) commitEditing();
    store.setSel([]);
    renderSelection();
    renderInspector();
    startMarquee(e);
  });
  slideHostEl.addEventListener('dblclick', onCanvasDblClick);
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
    zoomOut,
    zoomFit,
    zoomLabelEl,
    zoomIn
  ]);

  // Presentation layer
  presentHostEl = h('div', { class: 'slide-host' });
  presentEl = h('div', { class: 'present', onclick: presentNext }, [presentHostEl]);

  docInfoEl = h('div', { class: 'doc-info' });
  toastEl = h('div', { class: 'toast' });

  // Right-click context menu (on canvas)
  ctxMenuEl = h('div', { class: 'ctx-menu' });
  canvasArea.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    showCtxMenu(e.clientX, e.clientY);
  });
  document.addEventListener('click', () => hideCtxMenu());
  window.addEventListener('scroll', () => hideCtxMenu(), true);

  root.append(main, statusbar, docInfoEl, presentEl, toastEl, ctxMenuEl);

  document.addEventListener('keydown', onKeyDown);

  // Re-fit on viewport changes so the slide never overflows a small window (the default zoom is
  // "fit to window", matching the parser's preview & editor). Presentation mode re-computes its
  // own scale; otherwise only re-fit when the user hasn't picked a manual zoom.
  window.addEventListener('resize', onWindowResize);
}

/** Window resize handler: re-fit the canvas, or recompute the presentation scale if presenting. */
function onWindowResize(): void {
  if (presentEl.classList.contains('on')) {
    renderPresent();
  } else if (state.userZoom == null) {
    setZoom(null);
  }
}

// ---------- Preview pipeline (pptxToHtml: most faithful rendering) ----------
/**
 * Inject the parser's global stylesheet. pptxToHtml emits classed markup (_css_1 /
 * _tbl_cell_css_3, …) whose layout comes entirely from `styles.global`; without it the slides
 * render as unstyled HTML. Kept in its own <style> so the editor's own CSS stays separate.
 */
let globalStyleEl: HTMLStyleElement | null = null;
function ensureGlobalStyles(css: string | undefined): void {
  if (!css) return;
  if (!globalStyleEl) {
    globalStyleEl = document.createElement('style');
    globalStyleEl.id = 'pptx-global-styles';
    document.head.appendChild(globalStyleEl);
  }
  if (globalStyleEl.textContent !== css) globalStyleEl.textContent = css;
}

async function renderPreview(): Promise<void> {
  const bytes = state.latest;
  if (!bytes || !bytes.length) return;
  const result = await pptxToHtml(bytes, PARSE_OPTS);
  ensureGlobalStyles(result.styles?.global);
  state.preview = {
    slides: result.slides.map((s: any) => s.html),
    charts: result.charts || [],
    metadata: result.metadata,
    customProps: result.customProps
  };
  state.slideSize = result.slideSize || state.slideSize;
  if (state.present && state.present.version === state.version) state.present = null;
  if (state.mode === 'preview') {
    state.current = clamp(state.current, 0, state.preview.slides.length - 1);
    renderAll();
    fitIfAuto();
  }
}

// ---------- Edit pipeline (editor document -> DOM) ----------
async function enterEditMode(): Promise<void> {
  const bytes = state.latest;
  if (!bytes || !bytes.length) return;
  const standard = await pptxToStandard(bytes, PARSE_OPTS);
  const doc = docFromPptx(standard, { fileName: state.title });
  store.setDoc(doc, { noHistory: true });
  store.setSlide(state.current);
  state.slideSize = doc.slideSize || state.slideSize;
  state.current = clamp(state.current, 0, store.slideCount - 1);
  renderAll();
  fitIfAuto();
}

/** Serialize the edited document back to PPTX bytes. */
async function serializeDoc(): Promise<Uint8Array> {
  const out = await jsonToPptx(docToPptx(store.doc), { outputType: 'uint8array' });
  return out as Uint8Array;
}

/**
 * Push the current document to the host so it can save. Called after every edit (debounced) and
 * before anything that needs the up-to-date bytes (presenting, leaving edit mode).
 */
let syncQueued = false;
async function syncBytes(dirty: boolean): Promise<void> {
  if (state.syncing) {
    // Queue one follow-up pass instead of dropping it: edits made while a serialization is in
    // flight would otherwise never reach the host, and would silently be missing from the save.
    syncQueued = true;
    return;
  }
  state.syncing = true;
  try {
    if (!dirty) {
      // Undo landed back on the imported state: fall back to the original bytes
      state.latest = state.original;
      post({ type: 'sync', bytes: null, dirty: false });
      return;
    }
    state.latest = await serializeDoc();
    post({ type: 'sync', bytes: toBase64(state.latest), dirty: true });
  } catch (e: any) {
    post({ type: 'error', message: 'Cannot serialize this document: ' + (e?.message || String(e)) });
  } finally {
    state.syncing = false;
    if (syncQueued) {
      syncQueued = false;
      void syncBytes(dirty);
    }
  }
}
const syncSoon = debounce(() => syncBytes(store.canUndo()), 350);

/** Flush the latest edits to the host, then ask it to write the file. */
async function flushAndSave() {
  if (state.mode !== 'edit') return;
  if (store.editingId) commitEditing();
  await syncBytes(store.canUndo());
  post({ type: 'save' });
}

/** Run a store action and refresh everything that depends on the model. */
function runAction(fn: () => void, opts: { skipSync?: boolean } = {}) {
  if (state.mode !== 'edit') return;
  try {
    fn();
  } catch (e: any) {
    toast(e?.message || String(e), 'error');
    return;
  }
  state.version++;
  renderAll();
  if (!opts.skipSync) syncSoon();
}

// ---------- Mode switching ----------
async function toggleMode() {
  const next: EditorMode = state.mode === 'preview' ? 'edit' : 'preview';
  state.mode = next;
  post({ type: 'setMode', mode: next });
  applyMode();
}

async function applyMode() {
  root.className = 'app mode-' + state.mode;
  const editing = state.mode === 'edit';
  modeBtn.textContent = editing ? 'Preview' : 'Edit';
  modeBtn.title = editing ? 'Switch to preview mode' : 'Switch to edit mode';
  overlayEl.innerHTML = '';

  if (editing) {
    await enterEditMode();
    return;
  }
  // Leaving edit mode: flush the edited bytes first, then render them faithfully
  if (store.canUndo()) await syncBytes(true);
  disposeAllCharts();
  try {
    await renderPreview();
  } catch (e: any) {
    toast('Render failed: ' + (e?.message || String(e)), 'error');
  }
}

// ---------- Zoom ----------
/**
 * Re-fit to the window only when the user hasn't set an explicit zoom. This keeps a manual zoom
 * sticky across byte updates / mode switches / window resizes — the default zoom is "fit to
 * window" (like the parser's preview & editor), not a fixed 100%.
 */
function fitIfAuto(): void {
  if (state.userZoom == null) setZoom(null);
}

/** Fit the slide into the canvas area on both axes (capped at 2x). */
function getFitZoom(): number {
  const sc = document.getElementById('canvasScroll');
  const availW = (sc?.clientWidth || 800) - 40;
  const availH = (sc?.clientHeight || 600) - 40;
  return clamp(Math.min(availW / state.slideSize.width, availH / state.slideSize.height), 0.1, 2);
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

// ---------- Charts ----------
/**
 * Paint the ECharts placeholders emitted by pptxToHtml (preview + presentation). In edit mode
 * charts are rendered from the model by render.ts instead.
 */
function paintCharts(host: HTMLElement, charts?: any[]) {
  const list0 = charts ?? state.preview?.charts;
  if (!list0 || !list0.length) return;
  if (typeof (window as any).echarts === 'undefined') return; // skip silently if echarts not loaded
  try {
    // Container scope, so repeated copies of the same slide HTML (thumbnail / canvas / presentation)
    // never resolve to each other's chart containers.
    const list = list0.filter((c: any) => !!host.querySelector(`[id="${c.chartId}"]`));
    if (list.length) chartRenderer.renderCharts(list, host);
  } catch (e: any) {
    console.warn('[pptx-webview] Chart rendering failed:', e);
  }
}

/** Charts of the currently presented slide (rendered from their own render cache). */
function paintPresentCharts() {
  if (!state.present) return;
  paintCharts(presentHostEl, state.present.charts);
}

/**
 * Two ECharts instance registries exist: chartRenderer's singleton (used for pptxToHtml output)
 * and the renderer module's own map (used for model-rendered charts). Both hold instances whose
 * DOM has been replaced, so drop them to avoid leaks.
 */
function disposeStaleCharts() {
  disposeDetachedCharts();
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

// ---------- Rendering ----------
function renderAll() {
  applyStageTransform();
  gridEl.className = 'grid-overlay' + (state.grid ? ' on' : '');
  renderSlideList();
  renderCanvas();
  renderInspector();
  statusCountEl.textContent = `${slideCount()} slides`;
  undoBtn.disabled = state.mode !== 'edit' || !store.canUndo();
  redoBtn.disabled = state.mode !== 'edit' || !store.canRedo();
  if (docInfoEl.classList.contains('on')) renderDocInfo();
}

function slideCount(): number {
  return state.mode === 'edit' ? store.slideCount : state.preview?.slides.length || 0;
}

function renderSlideList() {
  slideListEl.innerHTML = '';
  const count = slideCount();
  for (let i = 0; i < count; i++) {
    const thumb = h('div', { class: 'thumb' + (i === state.current ? ' active' : ''), onclick: () => selectSlide(i) });
    if (state.mode === 'edit') {
      renderThumbInto(thumb, store.doc.slides[i], store.doc);
    } else {
      const htmls = state.preview?.slides || [];
      const scale = (thumb.clientWidth || 160) / state.slideSize.width;
      const inner = h('div', { class: 'inner' });
      inner.style.width = state.slideSize.width + 'px';
      inner.style.height = state.slideSize.height + 'px';
      inner.style.transform = `scale(${scale})`;
      inner.innerHTML = htmls[i] || '';
      thumb.style.height = Math.round(state.slideSize.height * scale) + 'px';
      thumb.appendChild(inner);
    }
    thumb.append(h('span', { class: 'num' }, [String(i + 1)]));
    slideListEl.append(thumb);
  }
}

function renderCanvas() {
  disposeStaleCharts();
  if (state.mode === 'edit') {
    const slide = store.slide;
    if (!slide) return;
    renderSlideInto(slideHostEl, slide, store.doc, { grid: state.grid, editingId: store.editingId });
    buildOverlay();
  } else {
    slideHostEl.innerHTML = state.preview?.slides[state.current] || '';
    paintCharts(slideHostEl);
    overlayEl.innerHTML = '';
  }
}

// ---------- Selection overlay (geometry from the parser core) ----------
function buildOverlay() {
  overlayEl.innerHTML = '';
  const slide = store.slide;
  if (!slide) return;
  for (const el of slide.elements || []) {
    if (el.hidden) continue;
    const r = elementRect(el);
    const m = effectMargin(el);
    const selected = store.sel.includes(el.id);
    const rect = h('div', {
      class: 'sel-rect' + (selected ? ' selected' : '') + (el.locked ? ' locked' : ''),
      'data-id': el.id,
      style:
        `left:${r.x - m.left}px;top:${r.y - m.top}px;` +
        `width:${r.width + m.left + m.right}px;height:${r.height + m.top + m.bottom}px;` +
        `transform:${el.rotation ? `rotate(${el.rotation}deg)` : ''}`
    });
    if (!selected) {
      rect.style.background = 'transparent';
      rect.style.borderColor = 'rgba(66,133,244,.5)';
      rect.style.pointerEvents = 'none';
    } else {
      rect.addEventListener('pointerdown', (e) => onRectPointerDown(e, el));
      if (!el.locked) {
        for (const d of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
          rect.append(h('div', { class: 'handle ' + d, 'data-dir': d }));
        }
        const rot = h('div', { class: 'handle rot', 'data-dir': 'rot', title: 'Rotate (hold Shift to snap 15°)' });
        rot.style.top = '-24px';
        rect.append(rot);
      }
    }
    if (el.locked) rect.append(h('span', { class: 'lock-badge' }, ['🔒']));
    overlayEl.append(rect);
  }
}

/** Cheap redraw used while the selection changes (keeps the canvas DOM untouched). */
function renderSelection() {
  buildOverlay();
  // Reflect the selection on the rendered elements too (used to unlock media controls)
  slideHostEl.querySelectorAll('.el[data-id]').forEach((node) => {
    node.classList.toggle('is-sel', store.sel.includes((node as HTMLElement).dataset.id!));
  });
}

// ---------- Drag / Resize / Rotate ----------
function onRectPointerDown(e: PointerEvent, el: any) {
  e.preventDefault();
  e.stopPropagation();
  const dir = (e.target as HTMLElement).getAttribute('data-dir');
  const startX = e.clientX;
  const startY = e.clientY;
  const orig = { x: num(el.x), y: num(el.y), w: num(el.width), h: num(el.height), rot: num(el.rotation) };
  const center = { x: orig.x + orig.w / 2, y: orig.y + orig.h / 2 };
  const startAngle = Math.atan2(e.clientY - center.y, e.clientX - center.x);
  const box = (e.currentTarget as HTMLElement);
  box.setPointerCapture(e.pointerId);
  // One history entry per gesture: snapshot up-front, then apply without history
  store.snapshot();

  const apply = (patch: any) => {
    store.update((doc: any) => {
      const t = findInDoc(doc, el.id);
      if (!t) return;
      Object.assign(t, patch);
    }, { history: false });
    const target: any = store.findElement(el.id);
    if (!target) return;
    const r = elementRect(target);
    const m = effectMargin(target);
    box.style.left = r.x - m.left + 'px';
    box.style.top = r.y - m.top + 'px';
    box.style.width = r.width + m.left + m.right + 'px';
    box.style.height = r.height + m.top + m.bottom + 'px';
    box.style.transform = target.rotation ? `rotate(${target.rotation}deg)` : '';
    // Keep the rendered element itself in step with the frame, otherwise only the box moves
    const node = slideHostEl.querySelector(`.el[data-id="${el.id}"]`) as HTMLElement | null;
    if (node) {
      node.style.left = `${r.x}px`;
      node.style.top = `${r.y}px`;
      node.style.width = `${r.width}px`;
      node.style.height = `${r.height}px`;
      node.style.transform = target.rotation ? `rotate(${target.rotation}deg)` : '';
    }
  };

  const onMove = (ev: PointerEvent) => {
    const dx = (ev.clientX - startX) / state.zoom;
    const dy = (ev.clientY - startY) / state.zoom;
    if (dir === 'rot') {
      // Rotate around the element centre; hold Shift to snap to 15°
      const ang = Math.atan2(ev.clientY - center.y, ev.clientX - center.x);
      let next = orig.rot + ((ang - startAngle) * 180) / Math.PI;
      if (ev.shiftKey) next = Math.round(next / 15) * 15;
      next = ((next + 180) % 360) - 180;
      apply({ rotation: Math.round(next * 10) / 10 });
      return;
    }
    if (!dir) {
      apply({ x: Math.round(orig.x + dx), y: Math.round(orig.y + dy) });
      return;
    }
    // Resize: rotate the pointer delta into the element's local frame so that rotated
    // elements still resize along their own axes
    const rad = (-orig.rot * Math.PI) / 180;
    const ldx = dx * Math.cos(rad) - dy * Math.sin(rad);
    const ldy = dx * Math.sin(rad) + dy * Math.cos(rad);
    const sx = dir.includes('w') ? -1 : dir.includes('e') ? 1 : 0;
    const sy = dir.includes('n') ? -1 : dir.includes('s') ? 1 : 0;
    let nw = orig.w;
    let nh = orig.h;
    if (sx) nw = Math.max(8, orig.w + sx * ldx);
    if (sy) nh = Math.max(8, orig.h + sy * ldy);
    if (ev.shiftKey && sx && sy) {
      const s = Math.max(nw / orig.w, nh / orig.h);
      nw = orig.w * s;
      nh = orig.h * s;
    }
    // The grabbed handle stays under the cursor: shift the centre by half the size delta,
    // then rotate that offset back into page coordinates.
    const lox = (sx * (nw - orig.w)) / 2;
    const loy = (sy * (nh - orig.h)) / 2;
    const rad2 = (orig.rot * Math.PI) / 180;
    const cx = orig.x + orig.w / 2 + lox * Math.cos(rad2) - loy * Math.sin(rad2);
    const cy = orig.y + orig.h / 2 + lox * Math.sin(rad2) + loy * Math.cos(rad2);
    apply({
      width: Math.round(nw),
      height: Math.round(nh),
      x: Math.round(cx - nw / 2),
      y: Math.round(cy - nh / 2)
    });
  };

  const onUp = (ev: PointerEvent) => {
    box.releasePointerCapture(ev.pointerId);
    box.removeEventListener('pointermove', onMove);
    box.removeEventListener('pointerup', onUp);
    state.version++;
    renderAll();
    syncSoon();
  };
  box.addEventListener('pointermove', onMove);
  box.addEventListener('pointerup', onUp);
}

// ---------- Canvas interaction: selection, marquee, smart guides, rich text ----------
/** Locate an element across the whole doc (top level or inside a group). */
function findInDoc(doc: any, id: string): any {
  for (const slide of doc.slides || []) {
    for (const el of slide.elements || []) {
      if (el.id === id) return el;
      for (const c of el.children || []) if (c.id === id) return c;
    }
  }
  return null;
}

/** Union bounding box of the given ids, in slide coordinates. */
function selUnion(ids: string[]) {
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (const id of ids) {
    const el = store.findElement(id);
    if (!el) continue;
    const r = elementRect(el);
    x0 = Math.min(x0, r.x);
    y0 = Math.min(y0, r.y);
    x1 = Math.max(x1, r.x + r.width);
    y1 = Math.max(y1, r.y + r.height);
  }
  return Number.isFinite(x0) ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null;
}

/** Move the rendered nodes with the model during a drag, instead of rebuilding the slide. */
function syncElementNodes(ids: string[]) {
  for (const id of ids) {
    const el = store.findElement(id);
    const node = slideHostEl.querySelector(`.el[data-id="${id}"]`) as HTMLElement | null;
    if (!el || !node) continue;
    const r = elementRect(el);
    node.style.left = `${r.x}px`;
    node.style.top = `${r.y}px`;
    node.style.width = `${r.width}px`;
    node.style.height = `${r.height}px`;
    node.style.transform = el.rotation ? `rotate(${el.rotation}deg)` : '';
  }
  buildOverlay();
}

/** Update only the selection highlight, leaving the canvas DOM untouched. */
function syncSelClasses() {
  slideHostEl.querySelectorAll('.el[data-id]').forEach((node) => {
    node.classList.toggle('is-sel', store.sel.includes((node as HTMLElement).dataset.id!));
  });
}

/**
 * Smart guides: snap the moving box to the edges/centres of sibling elements and of the slide.
 * Returns the correction to add to the pointer delta, plus the guide lines to draw.
 */
function snapBox(box: any, others: any[]) {
  const tol = 6 / (state.zoom || 1);
  const size = store.doc.slideSize || { width: 1280, height: 720 };
  const tx = [box.x, box.x + box.width / 2, box.x + box.width];
  const ty = [box.y, box.y + box.height / 2, box.y + box.height];
  const candX = [0, size.width / 2, size.width];
  const candY = [0, size.height / 2, size.height];
  for (const o of others) {
    candX.push(o.x, o.x + o.width / 2, o.x + o.width);
    candY.push(o.y, o.y + o.height / 2, o.y + o.height);
  }
  let bestX: any = null;
  let bestY: any = null;
  for (const t of tx)
    for (const c of candX) {
      const d = c - t;
      if (Math.abs(d) <= tol && (!bestX || Math.abs(d) < Math.abs(bestX.d))) bestX = { d, pos: c };
    }
  for (const t of ty)
    for (const c of candY) {
      const d = c - t;
      if (Math.abs(d) <= tol && (!bestY || Math.abs(d) < Math.abs(bestY.d))) bestY = { d, pos: c };
    }
  const guides: any[] = [];
  if (bestX) guides.push({ type: 'v', pos: bestX.pos });
  if (bestY) guides.push({ type: 'h', pos: bestY.pos });
  return { dx: bestX ? bestX.d : 0, dy: bestY ? bestY.d : 0, guides };
}

function showGuides(guides: any[]) {
  clearGuides();
  for (const g of guides) {
    overlayEl.append(
      h('div', { class: `guide ${g.type}`, style: g.type === 'v' ? `left:${g.pos}px` : `top:${g.pos}px` })
    );
  }
}
function clearGuides() {
  overlayEl.querySelectorAll('.guide').forEach((n) => n.remove());
}

/** Drag every selected element together. One undo entry per gesture. */
function startMove(e: PointerEvent) {
  const ids = store.sel.filter((id: string) => {
    const el = store.findElement(id);
    return el && !el.locked;
  });
  if (!ids.length) return;
  e.preventDefault();
  const startX = e.clientX;
  const startY = e.clientY;
  const orig = ids.map((id) => {
    const el = store.findElement(id);
    return {
      id,
      x: num(el.x),
      y: num(el.y),
      kids: (el.children || []).map((c: any) => ({ id: c.id, x: num(c.x), y: num(c.y) }))
    };
  });
  const box0 = selUnion(ids);
  if (!box0) return;
  const others = (store.slide?.elements || [])
    .filter((n: any) => n && !ids.includes(n.id) && !n.hidden)
    .map((n: any) => elementRect(n));
  store.snapshot();
  let moved = false;

  const onMove = (ev: PointerEvent) => {
    let dx = (ev.clientX - startX) / state.zoom;
    let dy = (ev.clientY - startY) / state.zoom;
    if (!moved && Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5) return;
    moved = true;
    let guides: any[] = [];
    if (store.snap !== false) {
      const s = snapBox(
        { x: box0.x + dx, y: box0.y + dy, width: box0.width, height: box0.height },
        others
      );
      dx += s.dx;
      dy += s.dy;
      guides = s.guides;
    }
    store.update((doc: any) => {
      for (const o of orig) {
        const el = findInDoc(doc, o.id);
        if (!el) continue;
        el.x = Math.round(o.x + dx);
        el.y = Math.round(o.y + dy);
        for (const k of o.kids) {
          const c = (el.children || []).find((x: any) => x.id === k.id);
          if (c) {
            c.x = Math.round(k.x + dx);
            c.y = Math.round(k.y + dy);
          }
        }
      }
    }, { history: false });
    syncElementNodes(ids);
    showGuides(guides);
  };
  const onUp = () => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    clearGuides();
    if (moved) {
      state.version++;
      renderAll();
      syncSoon();
    }
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}

/** Rubber-band selection over the canvas. */
function startMarquee(e: PointerEvent) {
  e.preventDefault();
  const startX = e.clientX;
  const startY = e.clientY;
  const box = h('div', { class: 'marquee' });
  overlayEl.append(box);
  const onMove = (ev: PointerEvent) => {
    const x0 = Math.min(startX, ev.clientX);
    const y0 = Math.min(startY, ev.clientY);
    const x1 = Math.max(startX, ev.clientX);
    const y1 = Math.max(startY, ev.clientY);
    if (x1 - x0 < 3 && y1 - y0 < 3) return;
    const r = slideHostEl.getBoundingClientRect();
    box.style.left = `${(x0 - r.left) / state.zoom}px`;
    box.style.top = `${(y0 - r.top) / state.zoom}px`;
    box.style.width = `${(x1 - x0) / state.zoom}px`;
    box.style.height = `${(y1 - y0) / state.zoom}px`;
    const bx0 = (x0 - r.left) / state.zoom;
    const by0 = (y0 - r.top) / state.zoom;
    const bx1 = (x1 - r.left) / state.zoom;
    const by1 = (y1 - r.top) / state.zoom;
    const hit = (store.slide?.elements || [])
      .filter((el: any) => el && !el.locked && !el.hidden)
      .map((el: any) => ({ id: el.id, r: elementRect(el) }))
      .filter((t: any) => t.r.x < bx1 && t.r.x + t.r.width > bx0 && t.r.y < by1 && t.r.y + t.r.height > by0)
      .map((t: any) => t.id);
    if (hit.length !== store.sel.length || hit.some((id: string) => !store.sel.includes(id))) {
      store.setSel(hit);
      syncSelClasses();
    }
  };
  const onUp = () => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    box.remove();
    renderSelection();
    renderInspector();
  };
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}

// ---------- Inline rich-text editing ----------
function richBodyNode(id: string): HTMLElement | null {
  return slideHostEl.querySelector(`.el[data-id="${id}"] [contenteditable="true"]`) as HTMLElement | null;
}

/** Double click: enter rich-text editing, or hand over to the type's own affordance. */
function onCanvasDblClick(e: MouseEvent) {
  if (state.mode !== 'edit') return;
  const node = (e.target as HTMLElement).closest('.el[data-id]') as HTMLElement | null;
  if (!node) return;
  const el = store.findElement(node.dataset.id!);
  if (!el || el.locked) return;
  store.setSel([el.id]);
  renderSelection();
  renderInspector();
  if (el.type === 'text') enterTextEditing(el);
  else if (el.type === 'image') pickImage();
  else if (el.type === 'table') editTable(el.id);
  else if (el.type === 'chart') editChart(el.id);
  else if (el.type === 'video' || el.type === 'audio') toast('Media cannot be played while editing');
  else toast(`Double-click editing is not supported for ${el.type} yet`);
}

/** Open the table editor against the freshest model state. */
function editTable(id: string) {
  const el = store.findElement(id);
  if (!el) return;
  openTableDialog(el, {
    update: (patch: any) => runAction(() => actions.updateElement(id, patch)),
    resize: (rows: number, cols: number) => runAction(() => actions.resizeTable(el, rows, cols)),
    reopen: () => editTable(id)
  });
}

/** Open the chart data editor against the freshest model state. */
function editChart(id: string) {
  const el = store.findElement(id);
  if (!el) return;
  openChartDialog(el, { update: (patch: any) => runAction(() => actions.updateElement(id, patch)) });
}

/** Turn one text element into a contenteditable body and keep it focused. */
function enterTextEditing(el: any) {
  if (store.editingId === el.id) return;
  store.editingId = el.id;
  state.editingText = true;
  renderAll();
  const body = richBodyNode(el.id);
  if (!body) return;
  body.focus();
  const range = document.createRange();
  range.selectNodeContents(body);
  range.collapse(false);
  const sel = window.getSelection();
  sel?.removeAllRanges();
  sel?.addRange(range);
  let timer: any = null;
  body.addEventListener('keydown', (ev: KeyboardEvent) => {
    ev.stopPropagation(); // keep the global shortcut bus out of the rich-text session
    if (ev.key === 'Escape') {
      ev.preventDefault();
      (ev.target as HTMLElement).blur();
    }
  });
  body.addEventListener('blur', () => {
    clearTimeout(timer);
    commitEditing();
  });
  body.addEventListener('input', () => {
    clearTimeout(timer);
    timer = setTimeout(() => commitEditing(true), 250);
  });
}

/** Write the contenteditable content back into the model as paragraphs/runs. */
function commitEditing(keepEditing = false) {
  const id = store.editingId;
  const el = id ? store.findElement(id) : null;
  const body = id ? richBodyNode(id) : null;
  if (!id || !el || !body) {
    store.editingId = null;
    state.editingText = false;
    return;
  }
  const paras = parseRichBody(body, el);
  const prev = Array.isArray(el.paragraphs) ? el.paragraphs : [];
  store.update((doc: any) => {
    const t = findInDoc(doc, id);
    if (!t) return;
    t.paragraphs = paras.map((p: any, i: number) => ({
      runs: p.runs,
      align: (prev[i] && prev[i].align) || el.align || 'left',
      bullet: (prev[i] && prev[i].bullet) || undefined,
      lineSpacing: (prev[i] && prev[i].lineSpacing) || el.lineSpacing || 1.15
    }));
  }, { coalesce: 'text:' + id });
  if (!keepEditing) {
    store.editingId = null;
    state.editingText = false;
    renderAll();
    syncSoon();
  }
}

/** contenteditable DOM -> model paragraphs/runs. Computed px styles are converted back to pt. */
function parseRichBody(body: HTMLElement, el: any) {
  const blocks = Array.from(body.querySelectorAll('div, p')) as HTMLElement[];
  const roots = blocks.length ? blocks : [body];
  const out: any[] = [];
  for (const block of roots) {
    const runs: any[] = [];
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let n: Node | null;
    while ((n = walker.nextNode())) {
      const text = n.nodeValue ?? '';
      if (!text) continue;
      const parent = (n.parentElement as HTMLElement) || block;
      // Bullet markers are decoration rendered by the view, not part of the text
      if (parent.classList?.contains('bullet-mark') || parent.classList?.contains('bullet-image')) continue;
      const cs = getComputedStyle(parent);
      const run: any = { text };
      if (/^(bold|bolder)$/i.test(cs.fontWeight) || parseInt(cs.fontWeight, 10) >= 600) run.bold = true;
      if (cs.fontStyle === 'italic' || cs.fontStyle === 'oblique') run.italic = true;
      if (cs.textDecorationLine.includes('underline') || cs.textDecoration.includes('underline')) run.underline = true;
      const size = parseFloat(cs.fontSize);
      if (Number.isFinite(size)) run.fontSize = Math.round((size * 72) / 96);
      if (cs.color) run.color = rgbToHex(cs.color);
      const ff = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (ff) run.fontFace = ff;
      const last = runs[runs.length - 1];
      if (last && sameRunStyle(last, run)) last.text += run.text;
      else runs.push(run);
    }
    out.push({ runs: runs.length ? runs : [{ text: '' }] });
  }
  return out;
}

function sameRunStyle(a: any, b: any) {
  const keys = ['bold', 'italic', 'underline', 'fontSize', 'color', 'fontFace'];
  return keys.every((k) => (a[k] ?? undefined) === (b[k] ?? undefined));
}

function rgbToHex(css: string): string {
  const m = css.match(/rgba?\(([^)]+)\)/);
  if (!m) return css;
  const parts = m[1].split(',').map((v) => parseInt(v, 10));
  return '#' + parts.slice(0, 3).map((v) => clamp(v, 0, 255).toString(16).padStart(2, '0')).join('');
}

// ---------- Inspector ----------
function renderInspector() {
  if (state.mode !== 'edit') {
    inspectorEl.innerHTML = '';
    return;
  }
  if (state.editingText && store.sel.length) return; // don't rebuild while typing (keeps focus)
  inspectorEl.innerHTML = '';
  const slide = store.slide;
  const count = store.slideCount;
  inspectorEl.append(h('h3', {}, ['Slide ' + (state.current + 1) + ' / ' + count]));

  // Element list
  const list = h('ul', { class: 'el-list' });
  for (const el of slide?.elements || []) {
    const tag = elSnippet(el);
    list.append(
      h(
        'li',
        {
          class: store.sel.includes(el.id) ? 'active' : '',
          onclick: () => {
            store.setSel([el.id]);
            renderSelection();
            renderInspector();
          }
        },
        [h('span', {}, [tag.name]), h('span', { class: 'tag' }, [tag.type])]
      )
    );
  }
  inspectorEl.append(h('h4', {}, ['Elements (click to select)']), list);

  const selected = store.sel.length === 1 ? store.findElement(store.sel[0]) : null;
  if (!selected) renderSlideInspector(slide);
  else renderElementInspector(selected);
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
  inspectorEl.append(h('h4', {}, ['Slide Properties']));

  inspectorEl.append(
    selectField(
      'Theme',
      THEMES.map((t: any) => t.id),
      store.doc.theme || 'blue',
      (v) => runAction(() => actions.applyTheme(v, true))
    )
  );
  inspectorEl.append(
    selectField(
      'Size',
      Object.keys(SLIDE_SIZES),
      store.doc.sizeKey || '16:9',
      (v) => runAction(() => actions.setSlideSize(SLIDE_SIZES[v]))
    )
  );
  inspectorEl.append(
    selectField(
      'Transition',
      TRANSITIONS.map((t: any) => t.value),
      slide?.transition?.type || 'none',
      (v) => runAction(() => actions.setTransition({ type: v, duration: 800, advanceOnClick: true }))
    )
  );
  inspectorEl.append(
    selectField(
      'Layout',
      LAYOUTS.map((l: any) => l.id),
      slide?.layout || 'titleBody',
      (v) => {
        // Applying a layout rebuilds the slide, so it replaces the current elements
        if (window.confirm('Applying a layout replaces every element on this slide. Continue?')) {
          runAction(() => actions.applyLayout(v));
        } else {
          renderInspector();
        }
      }
    )
  );

  inspectorEl.append(
    colorField(
      'Background',
      slide?.background && typeof slide.background === 'string' ? slide.background : '#ffffff',
      (v) => runAction(() => actions.setBackground(v))
    )
  );

  const hidden = h('input', { type: 'checkbox', ...(slide?.hidden ? { checked: 'checked' } : {}) });
  hidden.addEventListener('change', () =>
    runAction(() => store.update((doc: any) => {
      doc.slides[store.slideIndex].hidden = hidden.checked;
    }))
  );
  inspectorEl.append(h('div', { class: 'field' }, [h('label', {}, ['Hidden']), hidden]));

  const notes = h('textarea', { placeholder: 'Speaker notes...' }, [slide?.notes || '']);
  notes.addEventListener('focus', () => (state.editingText = true));
  notes.addEventListener('blur', () => (state.editingText = false));
  notes.addEventListener(
    'input',
    debounce(() => runAction(() => actions.setNotes(notes.value), { skipSync: true }), 400)
  );
  inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, ['Notes']), notes]));

  // ---- Animations on this slide ----
  inspectorEl.append(h('h4', {}, ['Animations']));
  const anims: any[] = Array.isArray(slide?.animations) ? slide.animations : [];
  if (!anims.length) inspectorEl.append(h('div', { class: 'empty' }, ['No animations on this slide']));
  for (let i = 0; i < anims.length; i++) {
    const a = anims[i];
    inspectorEl.append(
      h('div', { class: 'row' }, [
        h('span', { style: 'flex:1;font-size:11px' }, [`${a.presetClass || 'entr'} / ${a.type || '?'} / ${a.trigger || 'onClick'}`]),
        h('button', { class: 'btn', title: 'Move up', onclick: () => runAction(() => actions.moveAnimation(i, -1)) }, ['↑']),
        h('button', { class: 'btn', title: 'Move down', onclick: () => runAction(() => actions.moveAnimation(i, 1)) }, ['↓']),
        h('button', { class: 'btn', title: 'Remove', onclick: () => runAction(() => actions.removeAnimation(i)) }, ['✕'])
      ])
    );
  }
  // Adding an animation needs a target element, so that form lives in the element panel
  inspectorEl.append(h('div', { class: 'empty' }, ['Select an element to add animations for it.']));

  // Slide actions
  inspectorEl.append(h('h4', {}, ['Slide Actions']));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => addTextElement() }, ['+ Text Box']),
      h('button', { class: 'btn', onclick: () => addShapeElement('rect') }, ['+ Rectangle']),
      h('button', { class: 'btn', onclick: pickImage }, ['+ Image'])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => runAction(() => actions.duplicateSlide(state.current)) }, ['Duplicate']),
      h('button', {
        class: 'btn',
        onclick: () => {
          if (confirm('Delete this slide?')) runAction(() => actions.deleteSlide(state.current));
        }
      }, ['Delete'])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => runAction(() => actions.moveSlide(state.current, Math.max(0, state.current - 1))) }, ['Move Up']),
      h('button', { class: 'btn', onclick: () => runAction(() => actions.moveSlide(state.current, Math.min(store.slideCount - 1, state.current + 1))) }, ['Move Down'])
    ])
  );
}

/** The hyperlink shared by an element's runs ('' when none of them is linked). */
function firstHref(el: any): string {
  for (const p of el.paragraphs || []) for (const r of p.runs || []) if (r.href) return r.href;
  return '';
}

// ---- Element-level properties ----
function renderElementInspector(el: any) {
  inspectorEl.append(h('h4', {}, ['Element: ' + (el.type || 'unknown')]));

  const update = (patch: any) => runAction(() => actions.updateElement(el.id, patch));
  const geom = h('div', {});
  geom.append(numField('X', el.x, (v) => update({ x: v })));
  geom.append(numField('Y', el.y, (v) => update({ y: v })));
  geom.append(numField('Width', el.width, (v) => update({ width: Math.max(1, v) })));
  geom.append(numField('Height', el.height, (v) => update({ height: Math.max(1, v) })));
  geom.append(numField('Rotation', el.rotation, (v) => update({ rotation: v })));
  inspectorEl.append(geom);

  if (el.type === 'text') {
    const ta = h('textarea', { placeholder: 'Text content...' }, [el.text || '']);
    ta.addEventListener('focus', () => (state.editingText = true));
    ta.addEventListener('blur', () => (state.editingText = false));
    ta.addEventListener(
      'input',
      debounce(() => runAction(() => actions.updateElement(el.id, { text: ta.value }, { history: false })), 400)
    );
    inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, ['Text']), ta]));

    inspectorEl.append(colorField('Color', el.color || '#1e1e1e', (v) => update({ color: v })));
    inspectorEl.append(numField('Font Size', el.fontSize, (v) => update({ fontSize: v })));
    inspectorEl.append(selectField('Align', ['left', 'center', 'right', 'justify'], el.align || 'left', (v) => update({ align: v })));
    inspectorEl.append(selectField('Vertical', ['top', 'middle', 'bottom'], el.valign || 'top', (v) => update({ valign: v })));
    inspectorEl.append(checkField('Bold', !!el.bold, (v) => runAction(() => actions.applyTextStyleSel({ bold: v }))));
    inspectorEl.append(checkField('Italic', !!el.italic, (v) => runAction(() => actions.applyTextStyleSel({ italic: v }))));
    inspectorEl.append(checkField('Underline', !!el.underline, (v) => runAction(() => actions.applyTextStyleSel({ underline: v }))));
    inspectorEl.append(selectField('Font', FONT_LIST, el.fontFace || FONT_LIST[0], (v) => update({ fontFace: v })));
    inspectorEl.append(numField('Line Spacing', el.lineSpacing || 1.15, (v) => update({ lineSpacing: Math.max(0.5, v) })));
    inspectorEl.append(checkField('Bullet', !!el.bullet, (v) => runAction(() => actions.applyTextStyleSel({ bullet: !!v }))));
    // Hyperlinks live on runs, so writes are propagated to every run of the element
    inspectorEl.append(
      textField('Link', firstHref(el), (v) =>
        runAction(() => {
          const t = store.findElement(el.id);
          if (!t) return;
          const paras = (t.paragraphs || []).map((p: any) => ({
            ...p,
            runs: (p.runs || []).map((r: any) => ({ ...r, href: v || undefined }))
          }));
          actions.updateElement(el.id, { paragraphs: paras });
        })
      )
    );
  } else if (el.type === 'shape') {
    inspectorEl.append(textField('Shape', el.shapeType || 'rect', (v) => update({ shapeType: v })));
    const fillObj: any = el.fill && typeof el.fill === 'object' ? el.fill : null;
    const flatFill = fillObj ? fillObj.color || '#4285f4' : (typeof el.fill === 'string' && el.fill !== 'none' ? el.fill : '#4285f4');
    inspectorEl.append(
      selectField('Fill', ['solid', 'gradient', 'none'], fillObj?.type || 'solid', (v) => {
        if (v === 'none') update({ fill: 'none' });
        else if (v === 'gradient') update({ fill: { type: 'gradient', angle: 90, colors: [flatFill, '#ffffff'] } });
        else update({ fill: flatFill });
      })
    );
    if (!fillObj || fillObj.type !== 'none') {
      inspectorEl.append(
        colorField('Fill Color', flatFill, (v) => {
          if (fillObj && fillObj.type === 'gradient') update({ fill: { ...fillObj, colors: [v, fillObj.colors?.[1] || '#ffffff'] } });
          else update({ fill: v });
        })
      );
      if (fillObj && fillObj.type === 'gradient') {
        inspectorEl.append(
          colorField('Fill Color 2', fillObj.colors?.[1] || '#ffffff', (v) =>
            update({ fill: { ...fillObj, colors: [fillObj.colors?.[0] || flatFill, v] } })
          )
        );
      }
      inspectorEl.append(
        numField('Fill Transparency %', fillObj?.transparency ?? 0, (v) =>
          update({ fill: { ...(fillObj || { type: 'solid', color: flatFill }), transparency: clamp(v, 0, 100) } })
        )
      );
    }
    const lineColor = el.line && el.line !== 'none' ? el.line.color || '#000' : '#000';
    const lineWidth = el.line && el.line !== 'none' ? el.line.width || 1 : 1;
    inspectorEl.append(colorField('Border Color', lineColor, (v) => update({ line: { color: v, width: lineWidth } })));
    inspectorEl.append(numField('Border Width', lineWidth, (v) => update({ line: v > 0 ? { color: lineColor, width: v } : 'none' })));
  } else if (el.type === 'image') {
    const adj: any = el.imageAdjust || {};
    inspectorEl.append(h('div', { class: 'row' }, [h('button', { class: 'btn', onclick: pickImage }, ['Replace Image'])]));
    inspectorEl.append(numField('Brightness', adj.brightness ?? 0, (v) => update({ imageAdjust: { ...adj, brightness: clamp(v, -100, 100) } })));
    inspectorEl.append(numField('Contrast', adj.contrast ?? 0, (v) => update({ imageAdjust: { ...adj, contrast: clamp(v, -100, 100) } })));
    inspectorEl.append(numField('Transparency %', adj.transparency ?? 0, (v) => update({ imageAdjust: { ...adj, transparency: clamp(v, 0, 100) } })));
  } else if (el.type === 'table') {
    const rowCount = (el.rows || []).length;
    const colCount = Math.max(1, ...(el.rows || []).map((r: any) => (r.cells || []).length));
    inspectorEl.append(numField('Rows', rowCount, (v) => runAction(() => actions.resizeTable(el, Math.max(1, v), colCount))));
    inspectorEl.append(numField('Columns', colCount, (v) => runAction(() => actions.resizeTable(el, rowCount, Math.max(1, v)))));
    inspectorEl.append(checkField('Header Row', !!el.headerRow, (v) => update({ headerRow: v })));
    inspectorEl.append(colorField('Border Color', el.border?.color || '#cfd8e3', (v) => update({ border: { ...(el.border || { width: 1 }), color: v } })));
    inspectorEl.append(numField('Border Width', el.border?.width ?? 1, (v) => update({ border: { ...(el.border || { color: '#cfd8e3' }), width: v } })));
    inspectorEl.append(colorField('Header Fill', el.headerFill || '#1A73E8', (v) => update({ headerFill: v })));
    inspectorEl.append(colorField('Cell Fill', el.cellFill || '#ffffff', (v) => update({ cellFill: v })));
    inspectorEl.append(numField('Font Size', el.fontSize || 16, (v) => update({ fontSize: v })));
    inspectorEl.append(h('div', { class: 'row' }, [h('button', { class: 'btn', onclick: () => editTable(el.id) }, ['Edit Cells…'])]));
  } else if (el.type === 'chart') {
    inspectorEl.append(h('div', { class: 'row' }, [h('button', { class: 'btn', onclick: () => editChart(el.id) }, ['Edit Data…'])]));
    inspectorEl.append(h('div', { class: 'empty' }, ['Edit series and categories, or double-click the chart.']));
  } else if (el.type === 'video' || el.type === 'audio') {
    inspectorEl.append(textField('Name', el.name || '', (v) => update({ name: v })));
  } else {
    inspectorEl.append(h('div', { class: 'empty' }, ['This type only supports move / resize / delete']));
  }

  // Visual options shared by every element type
  inspectorEl.append(h('h4', {}, ['Appearance']));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => update({ flipH: !el.flipH }) }, [el.flipH ? 'Unflip H' : 'Flip H']),
      h('button', { class: 'btn', onclick: () => update({ flipV: !el.flipV }) }, [el.flipV ? 'Unflip V' : 'Flip V'])
    ])
  );
  inspectorEl.append(
    checkField('Shadow', !!el.shadow, (v) =>
      update({ shadow: v ? { type: 'outer', angle: 45, distance: 4, blur: 8, transparency: 60, color: '#000000' } : null })
    )
  );

  // Element actions
  inspectorEl.append(h('h4', {}, ['Element Actions']));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => runAction(() => actions.duplicateSelected()) }, ['Duplicate']),
      h('button', { class: 'btn', onclick: () => runAction(() => actions.deleteSelected()) }, ['Delete'])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => runAction(() => actions.zOrder('front')) }, ['Bring Front']),
      h('button', { class: 'btn', onclick: () => runAction(() => actions.zOrder('back')) }, ['Send Back'])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', title: 'Align left edge to slide', onclick: () => runAction(() => actions.alignElements('left')) }, ['Align L']),
      h('button', { class: 'btn', title: 'Align horizontal centre', onclick: () => runAction(() => actions.alignElements('hcenter')) }, ['Align C']),
      h('button', { class: 'btn', title: 'Align top edge to slide', onclick: () => runAction(() => actions.alignElements('top')) }, ['Align T']),
      h('button', { class: 'btn', title: 'Align vertical centre', onclick: () => runAction(() => actions.alignElements('vcenter')) }, ['Align M'])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', title: 'Align right edge', onclick: () => runAction(() => actions.alignElements('right')) }, ['Align R']),
      h('button', { class: 'btn', title: 'Align bottom edge', onclick: () => runAction(() => actions.alignElements('bottom')) }, ['Align B']),
      h('button', { class: 'btn', title: 'Distribute horizontally (needs 3+ selected)', onclick: () => runAction(() => actions.distribute('h')) }, ['Dist H']),
      h('button', { class: 'btn', title: 'Distribute vertically (needs 3+ selected)', onclick: () => runAction(() => actions.distribute('v')) }, ['Dist V'])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => runAction(() => actions.zOrder('forward')) }, ['Forward']),
      h('button', { class: 'btn', onclick: () => runAction(() => actions.zOrder('backward')) }, ['Backward']),
      h('button', { class: 'btn', onclick: () => runAction(() => actions.groupSelection()) }, ['Group']),
      h('button', { class: 'btn', onclick: () => runAction(() => actions.ungroupSelection()) }, ['Ungroup'])
    ])
  );
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', { class: 'btn', onclick: () => runAction(() => actions.toggleLock()) }, [el.locked ? 'Unlock' : 'Lock']),
      h('button', { class: 'btn', onclick: () => runAction(() => actions.toggleHidden()) }, [el.hidden ? 'Show' : 'Hide'])
    ])
  );

  // ---- Animations targeting this element ----
  inspectorEl.append(h('h4', {}, ['Animations']));
  const allAnims: any[] = Array.isArray(store.slide?.animations) ? store.slide.animations : [];
  const mine = allAnims.map((a: any, i: number) => ({ a, i })).filter((t: any) => t.a.target === el.id);
  if (!mine.length) inspectorEl.append(h('div', { class: 'empty' }, ['No animation on this element']));
  for (const t of mine) {
    inspectorEl.append(
      h('div', { class: 'row' }, [
        h('span', { style: 'flex:1;font-size:11px' }, [`${t.a.presetClass || 'entr'} / ${t.a.type} / ${t.a.trigger || 'onClick'}`]),
        h('button', { class: 'btn', title: 'Remove animation', onclick: () => runAction(() => actions.removeAnimation(t.i)) }, ['✕'])
      ])
    );
  }
  const pick = { cls: 'entr', type: (ANIM_TYPES.entr[0] || {}).value, trigger: 'onClick' };
  const typeWrap = h('div', {});
  const renderTypes = () => {
    typeWrap.innerHTML = '';
    const opts: string[] = (ANIM_TYPES[pick.cls] || []).map((x: any) => x.value);
    if (!opts.includes(pick.type)) pick.type = opts[0];
    typeWrap.append(selectField('Type', opts, pick.type, (v) => (pick.type = v)));
  };
  renderTypes();
  inspectorEl.append(
    selectField('Class', ANIM_CLASSES.map((c: any) => c.value), pick.cls, (v) => {
      pick.cls = v;
      renderTypes();
    })
  );
  inspectorEl.append(typeWrap);
  inspectorEl.append(selectField('Trigger', ['onClick', 'withPrev', 'afterPrev'], pick.trigger, (v) => (pick.trigger = v)));
  inspectorEl.append(
    h('div', { class: 'row' }, [
      h('button', {
        class: 'btn',
        onclick: () =>
          runAction(() =>
            actions.addAnimation({
              target: el.id,
              type: pick.type,
              presetClass: pick.cls,
              duration: 500,
              trigger: pick.trigger
            })
          )
      }, ['+ Animation'])
    ])
  );
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

// ---------- Slide selection ----------
function selectSlide(i: number) {
  state.current = clamp(i, 0, slideCount() - 1);
  if (state.mode === 'edit') store.setSlide(state.current);
  renderAll();
}

// ---------- Element creation ----------
/** New element id, matching the shape of the ids created while importing (`e_xxxx`). */
function newId(): string {
  return 'e_' + Math.random().toString(36).slice(2, 10);
}

function addTextElement() {
  runAction(() =>
    actions.addElement({
      id: newId(),
      type: 'text',
      x: 120,
      y: 120,
      width: 480,
      height: 90,
      text: 'Double-click to edit text',
      fontSize: 24,
      color: '#1e1e1e',
      align: 'left',
      valign: 'top'
    })
  );
}
function addShapeElement(shapeType: string) {
  runAction(() =>
    actions.addElement({ id: newId(), type: 'shape', shapeType, x: 200, y: 200, width: 200, height: 120, fill: '#4285f4', line: { color: '#000', width: 1 } })
  );
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
        const replaceId = store.sel.length === 1 && store.findElement(store.sel[0])?.type === 'image' ? store.sel[0] : null;
        runAction(() =>
          replaceId
            ? actions.updateElement(replaceId, { data: dataUrl })
            : actions.addElement({ id: newId(), type: 'image', data: dataUrl, x: 160, y: 160, width: w, height: hgt })
        );
      };
      img.src = dataUrl;
    };
    reader.readAsDataURL(file);
  });
  document.body.append(input);
  input.click();
}

function insertTable() {
  openTableSizeDialog({
    create: (rows: number, cols: number) => {
      const tblRows = Array.from({ length: rows }, () => ({
        cells: Array.from({ length: cols }, () => ({ text: '' }))
      }));
      runAction(() =>
        actions.addElement(
          {
            id: newId(),
            type: 'table',
            x: 160,
            y: 160,
            width: Math.max(240, cols * 90),
            height: Math.max(80, rows * 40),
            rows: tblRows,
            colWidths: new Array(cols).fill(1),
            headerRow: true,
            border: { color: '#cfd8e3', width: 1 },
            headerFill: '#1A73E8',
            fontSize: 16
          },
          { center: true }
        )
      );
    }
  });
}

function insertChart() {
  runAction(() =>
    actions.addElement(
      {
        id: newId(),
        type: 'chart',
        x: 160,
        y: 140,
        width: 520,
        height: 320,
        chartType: 'barChart',
        title: 'Chart title',
        legend: true,
        dataLabels: false,
        categories: ['A', 'B', 'C', 'D'],
        series: [{ name: 'Series 1', values: [5, 3, 8, 4] }]
      },
      { center: true }
    )
  );
}

function addMedia(type: 'video' | 'audio', data: string, ext: string) {
  runAction(() =>
    actions.addElement(
      {
        id: newId(),
        type,
        data,
        extension: ext,
        name: `${type}.${ext}`,
        x: 200,
        y: 160,
        width: type === 'video' ? 420 : 240,
        height: type === 'video' ? 260 : 90
      },
      { center: true }
    )
  );
}

// ---------- Grid / Presentation / Document info ----------
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
      h('tr', {}, [h('td', { class: 'k' }, [labels?.[k] || k]), h('td', { class: 'v' }, [String(obj[k])])])
    );
  }
  return table;
}

/** Render document properties: metadata (core.xml) + custom properties (custom.xml) */
function renderDocInfo() {
  docInfoEl.innerHTML = '';
  docInfoEl.append(
    h('div', { class: 'doc-info-head' }, [
      h('span', {}, ['Document Properties']),
      h('button', { class: 'btn', style: 'padding:2px 8px', onclick: () => docInfoEl.classList.remove('on') }, ['Close'])
    ])
  );
  const source = state.mode === 'edit' ? store.doc : state.preview;
  docInfoEl.append(h('h4', {}, ['Metadata']), kvTable(source?.metadata, META_LABELS));
  docInfoEl.append(h('h4', {}, ['Custom Properties']), kvTable(source?.customProps));
}

async function startPresent() {
  presentEl.className = 'present on';
  await renderPresent();
}
async function renderPresent() {
  // Present the most faithful rendering available: the current bytes through pptxToHtml
  disposeStaleCharts();
  if (!state.present || state.present.version !== state.version) {
    try {
      const bytes = state.latest;
      if (!bytes?.length) return;
      const result = await pptxToHtml(bytes, PARSE_OPTS);
      state.present = { slides: result.slides.map((s: any) => s.html), charts: result.charts || [], version: state.version };
    } catch (e: any) {
      toast('Present failed: ' + (e?.message || String(e)), 'error');
      return;
    }
  }
  const html = state.present.slides[state.current] || '';
  presentHostEl.style.width = state.slideSize.width + 'px';
  presentHostEl.style.height = state.slideSize.height + 'px';
  presentHostEl.innerHTML = html;
  // Charts need a measurable host, and media should only start once layout settled
  requestAnimationFrame(() => {
    if (state.present) paintPresentCharts();
    presentHostEl.querySelectorAll('video').forEach((v) => {
      v.play().catch(() => {
        /* autoplay may still be blocked; controls remain available */
      });
    });
  });
  const z = Math.min(window.innerWidth / state.slideSize.width, window.innerHeight / state.slideSize.height) * 0.96;
  presentHostEl.style.transform = `scale(${z})`;
}
function presentNext() {
  const total = state.present?.slides.length || slideCount();
  if (state.current < total - 1) {
    state.current++;
    renderPresent();
    if (state.mode === 'edit') store.setSlide(state.current);
    renderAll();
  } else {
    presentEl.className = 'present';
    disposeStaleCharts();
  }
}

// ---------- Context menu ----------
function showCtxMenu(x: number, y: number) {
  const items: { label: string; disabled?: boolean; onClick?: () => void }[] = [];
  if (state.mode === 'edit') {
    const hasSel = store.sel.length > 0;
    items.push(
      { label: 'Bring Front', disabled: !hasSel, onClick: () => runAction(() => actions.zOrder('front')) },
      { label: 'Send Back', disabled: !hasSel, onClick: () => runAction(() => actions.zOrder('back')) },
      { label: 'Duplicate', disabled: !hasSel, onClick: () => runAction(() => actions.duplicateSelected()) },
      { label: 'Delete', disabled: !hasSel, onClick: () => runAction(() => actions.deleteSelected()) },
      { label: '', onClick: () => {} },
      { label: store.selected().some((el: any) => el.locked) ? 'Unlock' : 'Lock', disabled: !hasSel, onClick: () => runAction(() => actions.toggleLock()) },
      { label: 'Select All', onClick: () => { actions.selectAll(); renderAll(); } },
      { label: '', onClick: () => {} },
      { label: 'Group', disabled: store.sel.length < 2, onClick: () => runAction(() => actions.groupSelection()) },
      { label: 'Ungroup', onClick: () => runAction(() => actions.ungroupSelection()) }
    );
  } else {
    items.push(
      { label: 'Edit', onClick: () => toggleMode() },
      { label: 'Document Properties', onClick: () => toggleDocInfo() }
    );
  }

  ctxMenuEl.innerHTML = '';
  for (const it of items) {
    if (!it.label) {
      ctxMenuEl.append(h('div', { class: 'ctx-sep' }));
      continue;
    }
    const node = h('div', { class: 'ctx-item' + (it.disabled ? ' disabled' : '') }, [it.label]);
    if (!it.disabled) node.addEventListener('click', () => { hideCtxMenu(); it.onClick?.(); });
    ctxMenuEl.append(node);
  }
  ctxMenuEl.classList.add('on');
  ctxMenuEl.style.left = Math.min(x, window.innerWidth - 200) + 'px';
  ctxMenuEl.style.top = Math.min(y, window.innerHeight - 300) + 'px';
}
function hideCtxMenu() {
  ctxMenuEl.classList.remove('on');
}

// ---------- Keyboard ----------
function onKeyDown(e: KeyboardEvent) {
  const target = e.target as HTMLElement;
  const typing = target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable);
  if (e.key === 'Escape' && hasOpenModal()) {
    closeModal();
    return;
  }
  if (e.key === 'Escape') {
    if (presentEl.classList.contains('on')) {
      presentEl.className = 'present';
      disposeStaleCharts();
      return;
    }
    if (state.mode === 'edit' && store.sel.length) {
      store.setSel([]);
      renderSelection();
      renderInspector();
    }
    return;
  }
  if (typing) return;
  if (presentEl.classList.contains('on')) {
    if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') presentNext();
    return;
  }
  if (state.mode !== 'edit') return;
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.key.toLowerCase() === 'z') {
    e.preventDefault();
    runAction(() => (e.shiftKey ? store.redo() : store.undo()));
    return;
  }
  if (mod && e.key.toLowerCase() === 'y') {
    e.preventDefault();
    runAction(() => store.redo());
    return;
  }
  if (mod && e.key.toLowerCase() === 'a') {
    e.preventDefault();
    actions.selectAll();
    renderAll();
    return;
  }
  if (mod && e.key.toLowerCase() === 'c') {
    e.preventDefault();
    actions.copySelected(false);
    renderAll();
    syncSoon();
    return;
  }
  if (mod && e.key.toLowerCase() === 'x') {
    e.preventDefault();
    actions.copySelected(true);
    renderAll();
    syncSoon();
    return;
  }
  if (mod && e.key.toLowerCase() === 'v') {
    e.preventDefault();
    actions.paste();
    renderAll();
    syncSoon();
    return;
  }
  if (mod && e.key.toLowerCase() === 'd') {
    e.preventDefault();
    runAction(() => actions.duplicateSelected());
    return;
  }
  if (mod && e.key.toLowerCase() === 'g') {
    e.preventDefault();
    runAction(() => (e.shiftKey ? actions.ungroupSelection() : actions.groupSelection()));
    return;
  }
  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    runAction(() => actions.deleteSelected());
    return;
  }
  if (e.key === 'Enter') {
    const one = store.sel.length === 1 ? store.findElement(store.sel[0]) : null;
    if (one && one.type === 'text') {
      e.preventDefault();
      enterTextEditing(one);
    }
    return;
  }
  if (e.key === 'Tab') {
    e.preventDefault();
    // Cycle through the elements of the current slide
    const els = (store.slide?.elements || []).filter((n: any) => n && !n.hidden);
    if (!els.length) return;
    const cur = store.sel.length ? els.findIndex((n: any) => n.id === store.sel[0]) : -1;
    const next = e.shiftKey ? (cur <= 0 ? els.length - 1 : cur - 1) : (cur + 1) % els.length;
    store.setSel([els[next].id]);
    renderSelection();
    renderInspector();
    return;
  }
  const nudges: Record<string, [number, number]> = {
    ArrowLeft: [-1, 0],
    ArrowRight: [1, 0],
    ArrowUp: [0, -1],
    ArrowDown: [0, 1]
  };
  const nudge = nudges[e.key];
  if (nudge && store.sel.length) {
    e.preventDefault();
    const step = e.shiftKey ? 10 : 1;
    runAction(() => actions.nudge(nudge[0] * step, nudge[1] * step));
  }
}

// ---------- Host messages ----------
window.addEventListener('message', async (event: MessageEvent<HostToWebview>) => {
  const msg = event.data;
  try {
    switch (msg.type) {
      case 'init': {
        state.original = toBytes(msg.originalBytes);
        state.latest = toBytes(msg.bytes);
        state.title = msg.title || state.title;
        state.mode = msg.mode === 'edit' ? 'edit' : 'preview';
        state.version++;
        await renderPreview();
        applyMode();
        break;
      }
      case 'bytes': {
        state.latest = toBytes(msg.bytes);
        state.version++;
        if (state.mode === 'preview') await renderPreview();
        break;
      }
      case 'saved':
        toast('Saved');
        break;
      case 'info':
        toast(msg.message, msg.kind || 'info');
        break;
    }
  } catch (e: any) {
    console.error('[pptx-webview] Message failed:', e);
    toast('Failed: ' + (e?.message || String(e)), 'error');
  }
});

// ---------- Bootstrap ----------
async function main() {
  const style = document.createElement('style');
  style.textContent = CSS;
  document.head.append(style);
  buildLayout();
  post({ type: 'ready' });
}
main().catch((e) => {
  console.error('[pptx-webview] Init failed:', e);
  document.body.textContent = 'Editor failed to start: ' + (e?.message || String(e));
});