// PPTX 编辑器 Webview 主逻辑
// 默认「预览模式」：直接用原始文件字节经 pptxToHtml 忠实渲染（等同 examples/index.html）。
// 「编辑模式」：切换到标准模型往返渲染，并叠加可选中/拖拽/改属性的编辑层。
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

// ---------- 全局状态 ----------
const state: {
  slideSize: { width: number; height: number };
  slidesHtml: string[]; // 当前展示的 HTML（预览或编辑，按需渲染后填充）
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

// 渲染源（宿主下发的字节）与渲染缓存：按需渲染，避免预览模式被迫等待模型往返
let previewSrc: any = null; // 原始文件字节 —— 忠实预览
let editSrc: any = null; // 模型往返后的字节 —— 编辑模式

/** 一次渲染结果：每页 HTML + 图表数据 + 文档信息（图表需由 echarts 二次绘制） */
interface RenderResult {
  slides: string[];
  charts: any[];
  metadata?: any;
  customProps?: Record<string, string>;
}
const rendered: { preview?: RenderResult; edit?: RenderResult } = {};

// ---------- DOM 引用 ----------
let root!: HTMLElement;
let slideListEl!: HTMLElement;
let stageEl!: HTMLElement;
let stageInnerEl!: HTMLElement;
let slideHostEl!: HTMLElement;
let overlayEl!: HTMLElement;
let gridEl!: HTMLElement;
let inspectorEl!: HTMLElement;
let titleEl!: HTMLElement;
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

// ---------- 工具函数 ----------
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

// ---------- 布局 ----------
function injectStyle() {
  const style = document.createElement('style');
  style.textContent = CSS;
  document.head.appendChild(style);
}

function buildLayout() {
  root = document.getElementById('app')!;
  root.className = 'app mode-preview';

  // 顶栏
  titleEl = h('div', { class: 'title' }, ['PPTX']);
  const saveBtn = h('button', { class: 'btn primary', title: '保存 (Ctrl+S)', onclick: () => post({ type: 'save' }) }, ['保存']);
  const saveAsBtn = h('button', { class: 'btn', title: '导出副本', onclick: () => post({ type: 'saveAs' }) }, ['导出']);
  const presentBtn = h('button', { class: 'btn', title: '演示 (F5)', onclick: startPresent }, ['演示']);
  const docBtn = h('button', { class: 'btn', title: '文档属性（元数据与自定义属性）', onclick: toggleDocInfo }, ['属性']);
  modeBtn = h('button', { class: 'btn', title: '切换到编辑模式', onclick: toggleMode }, ['编辑']);
  undoBtn = h('button', { class: 'btn edit-only', title: '撤销 (Ctrl+Z)', onclick: () => post({ type: 'undo' }) }, ['撤销']);
  redoBtn = h('button', { class: 'btn edit-only', title: '重做 (Ctrl+Y)', onclick: () => post({ type: 'redo' }) }, ['重做']);
  const topbar = h('div', { class: 'topbar' }, [
    titleEl,
    saveBtn,
    saveAsBtn,
    h('span', { class: 'sep' }),
    presentBtn,
    docBtn,
    h('span', { class: 'sep' }),
    modeBtn,
    undoBtn,
    redoBtn
  ]);

  // 幻灯片列表
  slideListEl = h('div', { class: 'sp-list' });
  const spHead = h('div', { class: 'sp-head' }, [
    h('span', {}, ['幻灯片']),
    h('button', {
      class: 'btn edit-only', title: '新建幻灯片', style: 'padding:1px 7px',
      onclick: () => sendOp({ kind: 'slideAdd', after: state.current })
    }, ['＋'])
  ]);
  const slidesPanel = h('div', { class: 'slides-panel' }, [spHead, slideListEl]);

  // 画布
  slideHostEl = h('div', { class: 'slide-host' });
  overlayEl = h('div', { class: 'overlay' });
  gridEl = h('div', { class: 'grid-overlay' });
  stageInnerEl = h('div', { class: 'stage-inner' }, [slideHostEl, gridEl, overlayEl]);
  stageEl = h('div', { class: 'stage' }, [stageInnerEl]);
  const scroll = h('div', { class: 'canvas-scroll', id: 'canvasScroll' }, [stageEl]);
  const canvasArea = h('div', { class: 'canvas-area' }, [scroll]);

  // 属性面板（仅编辑模式）
  inspectorEl = h('div', { class: 'inspector edit-only' });
  const main = h('div', { class: 'main' }, [slidesPanel, canvasArea, inspectorEl]);

  // 状态栏
  statusCountEl = h('span', {}, ['0 页']);
  zoomLabelEl = h('span', { class: 'chip', title: '点击恢复 100%', onclick: () => setZoom(1) }, ['100%']);
  const zoomOut = h('span', { class: 'chip', title: '缩小', onclick: () => setZoom(state.userZoom ?? state.zoom, true, 0.9) }, ['−']);
  const zoomIn = h('span', { class: 'chip', title: '放大', onclick: () => setZoom(state.userZoom ?? state.zoom, true, 1.1) }, ['＋']);
  const zoomFit = h('span', { class: 'chip', title: '适应窗口', onclick: () => setZoom(null) }, ['适应']);
  gridBtn = h('span', { class: 'chip edit-only', title: '网格', onclick: toggleGrid }, ['网格']);
  const statusbar = h('div', { class: 'statusbar' }, [
    statusCountEl,
    h('span', { class: 'spacer' }),
    gridBtn,
    zoomOut, zoomFit, zoomLabelEl, zoomIn
  ]);

  // 演示层
  presentHostEl = h('div', { class: 'slide-host' });
  presentEl = h('div', { class: 'present', onclick: presentNext }, [presentHostEl]);

  docInfoEl = h('div', { class: 'doc-info' });

  toastEl = h('div', { class: 'toast' });

  root.append(topbar, main, statusbar, docInfoEl, presentEl, toastEl);
}

// ---------- 模式切换 ----------
function toggleMode() {
  const next: 'preview' | 'edit' = state.mode === 'preview' ? 'edit' : 'preview';
  state.mode = next;
  state.selected = null;
  post({ type: 'setMode', mode: next });
  applyMode();
}
function applyMode() {
  root.className = 'app mode-' + state.mode;
  modeBtn.textContent = state.mode === 'preview' ? '编辑' : '预览';
  modeBtn.title = state.mode === 'preview' ? '切换到编辑模式' : '切换到预览模式';

  const htmls = rendered[state.mode]?.slides;
  if (htmls) {
    state.slidesHtml = htmls;
    state.current = clamp(state.current, 0, state.slidesHtml.length - 1);
    renderAll();
    return;
  }
  // 缓存未就绪：先渲染空态，渲染完成后再刷新（不阻塞界面）
  state.slidesHtml = [];
  renderAll();
  renderMode(state.mode);
}

/** 渲染当前模式；失败时通过 toast 提示，绝不静默吞掉 */
async function renderMode(mode: 'preview' | 'edit') {
  const src = mode === 'preview' ? previewSrc : editSrc;
  if (!src) return;
  try {
    // 该模式的字节尚未就绪时（如预览模式下宿主不做模型往返）静默跳过，不算失败
    if (!toBytes(src).length) return;
    rendered[mode] = await renderFromBytes(src, mode === 'preview');
    if (state.mode === mode) {
      state.slidesHtml = rendered[mode]!.slides;
      state.current = clamp(state.current, 0, state.slidesHtml.length - 1);
      renderAll();
    }
  } catch (e: any) {
    console.error('[pptx-webview] 渲染失败:', e);
    toast('渲染失败：' + (e?.message || String(e)), 'error');
  }
}

// ---------- 缩放 ----------
/** 适应宽度（与 examples/index.html 的 fitToWidth 一致，上限 2 倍） */
function getFitZoom(): number {
  const sc = document.getElementById('canvasScroll');
  const avail = (sc?.clientWidth || 800) - 40; // 减去左右 padding
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
  // 外层占位 = 缩放后尺寸，避免放大时被裁切、缩小后留有空白滚动区
  stageEl.style.width = Math.round(state.slideSize.width * state.zoom) + 'px';
  stageEl.style.height = Math.round(state.slideSize.height * state.zoom) + 'px';
}

// ---------- 图表（echarts） ----------
/**
 * 同一页 HTML 会出现在缩略图 / 画布 / 演示层，导致图表 id 重复，而
 * chart-renderer 用 document.getElementById 取容器（只会命中文档中的第一个）。
 * 因此给非主画布的副本加 id 前缀，保证主画布的 chart 容器唯一。
 */
function prefixChartIds(html: string, prefix: string): string {
  return html.replace(/id=(["'])chart([^"']*)\1/g, `id=$1${prefix}chart$2$1`);
}

/** 容器被替换后，旧的 echarts 实例已脱离文档，销毁以免泄漏 */
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

/** 用 echarts 绘制 host 内存在的图表（解析器只产出空的占位 div） */
function paintCharts(host: HTMLElement, idPrefix = '') {
  const r = rendered[state.mode];
  if (!r || !r.charts.length) return;
  if (typeof (window as any).echarts === 'undefined') return; // echarts 未加载时静默跳过
  const list = r.charts.filter((c) => !!host.querySelector('#' + idPrefix + c.chartId));
  if (!list.length) return;
  try {
    chartRenderer.renderCharts(
      idPrefix ? list.map((c) => ({ ...c, chartId: idPrefix + c.chartId })) : list
    );
  } catch (e: any) {
    console.warn('[pptx-webview] 图表渲染失败:', e);
  }
}

// ---------- 渲染 ----------
function renderAll() {
  renderSlideList();
  renderCanvas();
  renderInspector();
  statusCountEl.textContent = `${state.slidesHtml.length} 页`;
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
  // 预览模式不显示属性面板内容
  if (state.mode !== 'edit') {
    inspectorEl.innerHTML = '';
    return;
  }
  // 编辑文本时不重建，避免失焦
  if (state.editingText && state.selected) return;
  inspectorEl.innerHTML = '';
  const slide = state.model?.slides?.[state.current];

  inspectorEl.append(h('h3', {}, ['幻灯片 ' + (state.current + 1) + ' / ' + state.slidesHtml.length]));

  // 元素列表
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
  inspectorEl.append(h('h4', {}, ['元素 (点击选择)']), list);

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
      return { name: (el.text || '').toString().split('\n')[0].slice(0, 24) || '文本', type: '文本' };
    case 'shape':
      return { name: (el.shapeType || 'shape').toString(), type: '形状' };
    case 'image':
      return { name: '图片', type: '图片' };
    case 'chart':
      return { name: (el.chartType || 'chart').toString(), type: '图表' };
    case 'table':
      return { name: '表格', type: '表格' };
    case 'group':
      return { name: '组合', type: '组合' };
    case 'diagram':
      return { name: '图示', type: '图示' };
    case 'video':
      return { name: '视频', type: '视频' };
    case 'audio':
      return { name: '音频', type: '音频' };
    default:
      return { name: (el.type || 'element').toString(), type: el.type || '元素' };
  }
}

// ---- 幻灯片级属性 ----
function renderSlideInspector(slide: any) {
  if (!slide) return;
  inspectorEl.append(h('h4', {}, ['幻灯片属性']));

  const bg = colorField('背景', slide.background && typeof slide.background === 'string' ? slide.background : '#ffffff', (v) =>
    sendOp({ kind: 'slideUpdate', index: state.current, patch: { background: v } })
  );
  inspectorEl.append(bg);

  const hidden = h('input', { type: 'checkbox', ...(slide.hidden ? { checked: 'checked' } : {}) });
  hidden.addEventListener('change', () => sendOp({ kind: 'slideUpdate', index: state.current, patch: { hidden: hidden.checked } }));
  inspectorEl.append(h('div', { class: 'field' }, [h('label', {}, ['隐藏']), hidden]));

  const notes = h('textarea', { placeholder: '演讲者备注…' }, [slide.notes || '']);
  notes.addEventListener('input', debounce(() => sendOp({ kind: 'slideUpdate', index: state.current, patch: { notes: notes.value } }), 400));
  inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, ['备注']), notes]));

  // 幻灯片操作
  inspectorEl.append(h('h4', {}, ['幻灯片操作']));
  const addText = h('button', { class: 'btn', onclick: () => addTextElement() }, ['＋ 文本框']);
  const addShape = h('button', { class: 'btn', onclick: () => addShapeElement('rect') }, ['＋ 矩形']);
  const addImg = h('button', { class: 'btn', onclick: pickImage }, ['＋ 图片']);
  inspectorEl.append(h('div', { class: 'row' }, [addText, addShape, addImg]));

  const dup = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'slideDuplicate', index: state.current }) }, ['复制页']);
  const del = h('button', { class: 'btn', onclick: () => { if (confirm('删除当前幻灯片？')) sendOp({ kind: 'slideDelete', index: state.current }); } }, ['删除页']);
  const up = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'slideMove', from: state.current, to: Math.max(0, state.current - 1) }) }, ['上移']);
  const down = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'slideMove', from: state.current, to: Math.min(state.slidesHtml.length - 1, state.current + 1) }) }, ['下移']);
  inspectorEl.append(h('div', { class: 'row' }, [dup, del]));
  inspectorEl.append(h('div', { class: 'row' }, [up, down]));
}

// ---- 元素级属性 ----
function renderElementInspector(el: any) {
  inspectorEl.append(h('h4', {}, ['元素：' + (el.type || '未知')]));

  // 几何
  const geom = h('div', {});
  geom.append(numField('X', el.x, (v) => updateEl({ x: v })));
  geom.append(numField('Y', el.y, (v) => updateEl({ y: v })));
  geom.append(numField('宽', el.width, (v) => updateEl({ width: Math.max(1, v) })));
  geom.append(numField('高', el.height, (v) => updateEl({ height: Math.max(1, v) })));
  geom.append(numField('旋转', el.rotation, (v) => updateEl({ rotation: v })));
  inspectorEl.append(geom);

  // 类型相关
  if (el.type === 'text') {
    const ta = h('textarea', { placeholder: '文本内容…' }, [el.text || '']);
    ta.addEventListener('focus', () => (state.editingText = true));
    ta.addEventListener('blur', () => { state.editingText = false; });
    ta.addEventListener('input', debounce(() => { updateElLocal({ text: ta.value }); sendOp({ kind: 'elementUpdate', slide: state.current, element: state.selected!.element, patch: { text: ta.value } }); }, 400));
    inspectorEl.append(h('div', { class: 'field', style: 'align-items:flex-start' }, [h('label', {}, ['文本']), ta]));

    inspectorEl.append(colorField('颜色', el.color || '#1e1e1e', (v) => updateEl({ color: v })));
    inspectorEl.append(numField('字号', el.fontSize, (v) => updateEl({ fontSize: v })));
    inspectorEl.append(selectField('对齐', ['left', 'center', 'right', 'justify'], el.align || 'left', (v) => updateEl({ align: v })));
    inspectorEl.append(selectField('垂直', ['top', 'middle', 'bottom'], el.valign || 'top', (v) => updateEl({ valign: v })));
    inspectorEl.append(checkField('加粗', !!el.bold, (v) => updateEl({ bold: v })));
    inspectorEl.append(checkField('斜体', !!el.italic, (v) => updateEl({ italic: v })));
    inspectorEl.append(checkField('下划线', !!el.underline, (v) => updateEl({ underline: v })));
  } else if (el.type === 'shape') {
    inspectorEl.append(textField('形状', el.shapeType || 'rect', (v) => updateEl({ shapeType: v })));
    const fillVal = typeof el.fill === 'string' ? el.fill : (el.fill && el.fill.color) || '#4285f4';
    inspectorEl.append(colorField('填充', fillVal, (v) => updateEl({ fill: v })));
    const lineColor = el.line && el.line !== 'none' ? (el.line.color || '#000') : '#000';
    const lineWidth = el.line && el.line !== 'none' ? (el.line.width || 1) : 1;
    inspectorEl.append(colorField('边框色', lineColor, (v) => updateEl({ line: { color: v, width: lineWidth } })));
    inspectorEl.append(numField('边框宽', lineWidth, (v) => updateEl({ line: v > 0 ? { color: lineColor, width: v } : 'none' })));
  } else if (el.type === 'image') {
    const replace = h('button', { class: 'btn', onclick: pickImage }, ['替换图片']);
    inspectorEl.append(h('div', { class: 'row' }, [replace]));
    inspectorEl.append(h('div', { class: 'empty' }, ['图片内容来自预览，替换请选择本地文件']));
  } else {
    inspectorEl.append(h('div', { class: 'empty' }, ['该类型暂仅支持移动 / 缩放 / 删除']));
  }

  // 元素操作
  inspectorEl.append(h('h4', {}, ['元素操作']));
  const del = h('button', { class: 'btn', onclick: () => sendOp({ kind: 'elementDelete', slide: state.current, element: state.selected!.element }) }, ['删除']);
  const fwd = h('button', { class: 'btn', onclick: () => reorderEl(1) }, ['上移一层']);
  const bwd = h('button', { class: 'btn', onclick: () => reorderEl(-1) }, ['下移一层']);
  inspectorEl.append(h('div', { class: 'row' }, [del]));
  inspectorEl.append(h('div', { class: 'row' }, [fwd, bwd]));
}

// ---------- 字段构造 ----------
function colorField(label: string, value: string, onInput: (v: string) => void) {
  const input = h('input', { type: 'color', value: toHex(value) });
  input.addEventListener('input', () => onInput(input.value));
  const clear = h('button', { class: 'btn', style: 'padding:2px 6px', title: '设为无', onclick: () => onInput('none') }, ['无']);
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

// ---------- 选择 / 操作 ----------
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
    element: { type: 'text', x: 120, y: 120, width: 480, height: 90, text: '双击编辑文本', fontSize: 24, color: '#1e1e1e', align: 'left', valign: 'top' }
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

// ---------- 拖拽 / 缩放 ----------
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

// ---------- 网格 / 演示 ----------
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
  title: '标题',
  subject: '主题',
  author: '作者',
  keywords: '关键词',
  description: '备注',
  lastModifiedBy: '最后修改者',
  created: '创建时间',
  modified: '修改时间',
  category: '类别',
  status: '状态',
  contentType: '内容类型',
  language: '语言'
};

function kvTable(obj: any, labels?: Record<string, string>) {
  const keys = Object.keys(obj || {}).filter((k) => obj[k] != null && obj[k] !== '');
  if (!keys.length) return h('div', { class: 'empty' }, ['（无）']);
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

/** 渲染文档属性：元数据（core.xml）+ 自定义属性（custom.xml） */
function renderDocInfo() {
  const r = rendered[state.mode];
  docInfoEl.innerHTML = '';
  docInfoEl.append(
    h('div', { class: 'doc-info-head' }, [
      h('span', {}, ['文档属性']),
      h('button', { class: 'btn', style: 'padding:2px 8px', onclick: () => docInfoEl.classList.remove('on') }, ['关闭'])
    ])
  );
  docInfoEl.append(h('h4', {}, ['元数据']), kvTable(r?.metadata, META_LABELS));
  docInfoEl.append(h('h4', {}, ['自定义属性']), kvTable(r?.customProps));
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

// ---------- 消息处理 ----------
/** 把宿主传来的字节统一成 Uint8Array（兼容 TypedArray / ArrayBuffer / 数组 / base64 / 纯对象降级形态） */
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
    // postMessage 降级为纯对象时形如 {"0":80,"1":75,...}
    const keys = Object.keys(v).filter((k) => /^\d+$/.test(k));
    if (keys.length) {
      const out = new Uint8Array(keys.length);
      for (const k of keys) out[Number(k)] = Number(v[k]) & 0xff;
      return out;
    }
  }
  throw new Error('无法识别的 PPTX 字节数据');
}

/** 注入解析器产出的全局样式。缺少它会导致颜色/行高/SVG 滤镜全部丢失（渲染错乱） */
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
 * 用 pptxToHtml 把字节渲染成每页 HTML。
 * 与 examples/index.html 保持一致：注入 styles.global；预览模式跳过隐藏页（p:sld show="0"）。
 */
async function renderFromBytes(bytes: any, skipHidden = false): Promise<RenderResult> {
  const data = toBytes(bytes);
  if (!data.length) throw new Error('PPTX 字节为空，无法解析');
  const res = await pptxToHtml(data, { mediaProcess: true, themeProcess: true });
  if (!res) throw new Error('pptxToHtml 未返回解析结果');
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
  titleEl.textContent = msg.title;

  // 记录渲染源；只渲染当前模式（预览用原始字节忠实渲染，编辑用模型往返结果）
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

// ---------- 键盘 ----------
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
  if (state.mode !== 'edit') return; // 预览模式禁用编辑快捷键
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

// ---------- 启动 ----------
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
        // 不能让异常静默吞掉：否则界面会停在「0 页」而毫无提示
        console.error('[pptx-webview] 处理消息失败:', e);
        toast('渲染失败：' + (e?.message || String(e)), 'error');
      }
    })();
  });
  post({ type: 'ready' });
}

main();
