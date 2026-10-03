// 将一次编辑操作应用到标准 PptxDocument 模型（原地修改）。
// 模型结构见 @fefeding/ppt-parser 的 PptxDocument（types/pptx-document.ts）。
import type { PptxOp } from './protocol';

/** 深拷贝（模型为纯 JSON，无函数） */
export function cloneModel(model: any): any {
  return JSON.parse(JSON.stringify(model));
}

function blankSlide(): any {
  return { background: '#ffffff', elements: [] };
}

/** 应用一次编辑到模型（原地修改）。非法操作抛错，由调用方决定是否提交。 */
export function applyOpToModel(model: any, op: PptxOp): void {
  const slides: any[] = model.slides;
  if (!Array.isArray(slides)) throw new Error('模型缺少 slides 数组');

  switch (op.kind) {
    case 'slideAdd': {
      const idx = op.after == null ? slides.length : clamp(op.after, 0, slides.length);
      slides.splice(idx, 0, op.slide ? cloneModel(op.slide) : blankSlide());
      break;
    }
    case 'slideDelete': {
      if (slides.length <= 1) throw new Error('至少需保留一页幻灯片');
      if (op.index < 0 || op.index >= slides.length) throw new Error('页码越界');
      slides.splice(op.index, 1);
      break;
    }
    case 'slideDuplicate': {
      if (op.index < 0 || op.index >= slides.length) throw new Error('页码越界');
      slides.splice(op.index + 1, 0, cloneModel(slides[op.index]));
      break;
    }
    case 'slideMove': {
      if (op.from < 0 || op.from >= slides.length || op.to < 0 || op.to >= slides.length)
        throw new Error('页码越界');
      const [m] = slides.splice(op.from, 1);
      slides.splice(op.to, 0, m);
      break;
    }
    case 'slideUpdate': {
      const s = slides[op.index];
      if (!s) throw new Error('页码越界');
      Object.assign(s, op.patch);
      break;
    }
    case 'elementAdd': {
      const s = slides[op.slide];
      if (!s) throw new Error('页码越界');
      if (!Array.isArray(s.elements)) s.elements = [];
      s.elements.push(op.element);
      break;
    }
    case 'elementUpdate': {
      const s = slides[op.slide];
      if (!s || !Array.isArray(s.elements)) throw new Error('页码越界');
      const el = s.elements[op.element];
      if (!el) throw new Error('元素索引越界');
      Object.assign(el, op.patch);
      break;
    }
    case 'elementDelete': {
      const s = slides[op.slide];
      if (!s || !Array.isArray(s.elements)) throw new Error('页码越界');
      if (op.element < 0 || op.element >= s.elements.length) throw new Error('元素索引越界');
      s.elements.splice(op.element, 1);
      break;
    }
    case 'elementReorder': {
      const s = slides[op.slide];
      if (!s || !Array.isArray(s.elements)) throw new Error('页码越界');
      const arr = s.elements;
      if (op.from < 0 || op.from >= arr.length || op.to < 0 || op.to >= arr.length)
        throw new Error('元素索引越界');
      const [m] = arr.splice(op.from, 1);
      arr.splice(op.to, 0, m);
      break;
    }
    case 'metadataUpdate': {
      model.metadata = model.metadata || {};
      Object.assign(model.metadata, op.patch);
      break;
    }
    default: {
      const _exhaustive: never = op;
      throw new Error('未知操作: ' + JSON.stringify(_exhaustive));
    }
  }
}

function clamp(v: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, v));
}

/**
 * 裁剪模型用于发送到 Webview：移除大体积字段（图片 base64、__raw 原始子树），
 * 仅保留编辑所需的几何与文本信息，显著降低消息体积。
 */
export function trimModel(model: any): any {
  const clone = cloneModel(model);
  if (Array.isArray(clone.slides)) {
    for (const slide of clone.slides) {
      if (!slide || !Array.isArray(slide.elements)) continue;
      for (const el of slide.elements) {
        if (!el || typeof el !== 'object') continue;
        delete el.__raw;
        delete el.rawFallback;
        if (el.type === 'image' || el.type === 'video' || el.type === 'audio' || el.type === 'ole') {
          delete el.data; // 预览图来自 slidesHtml，无需在模型里重复下发
        }
      }
    }
  }
  if (clone.media) delete clone.media;
  if (clone.fonts) delete clone.fonts;
  return clone;
}
