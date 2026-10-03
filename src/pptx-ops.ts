// Applies a single edit operation to the standard PptxDocument model (in-place).
// Model structure: see @fefeding/ppt-parser PptxDocument (types/pptx-document.ts).
import type { PptxOp } from './protocol';

/** Deep clone (model is pure JSON, no functions). */
export function cloneModel(model: any): any {
  return JSON.parse(JSON.stringify(model));
}

function blankSlide(): any {
  return { background: '#ffffff', elements: [] };
}

/** Apply an edit to the model (in-place). Throws on invalid operations; caller decides whether to commit. */
export function applyOpToModel(model: any, op: PptxOp): void {
  const slides: any[] = model.slides;
  if (!Array.isArray(slides)) throw new Error('Model missing slides array');

  switch (op.kind) {
    case 'slideAdd': {
      const idx = op.after == null ? slides.length : clamp(op.after, 0, slides.length);
      slides.splice(idx, 0, op.slide ? cloneModel(op.slide) : blankSlide());
      break;
    }
    case 'slideDelete': {
      if (slides.length <= 1) throw new Error('At least one slide must remain');
      if (op.index < 0 || op.index >= slides.length) throw new Error('Slide index out of bounds');
      slides.splice(op.index, 1);
      break;
    }
    case 'slideDuplicate': {
      if (op.index < 0 || op.index >= slides.length) throw new Error('Slide index out of bounds');
      slides.splice(op.index + 1, 0, cloneModel(slides[op.index]));
      break;
    }
    case 'slideMove': {
      if (op.from < 0 || op.from >= slides.length || op.to < 0 || op.to >= slides.length)
        throw new Error('Slide index out of bounds');
      const [m] = slides.splice(op.from, 1);
      slides.splice(op.to, 0, m);
      break;
    }
    case 'slideUpdate': {
      const s = slides[op.index];
      if (!s) throw new Error('Slide index out of bounds');
      Object.assign(s, op.patch);
      break;
    }
    case 'elementAdd': {
      const s = slides[op.slide];
      if (!s) throw new Error('Slide index out of bounds');
      if (!Array.isArray(s.elements)) s.elements = [];
      s.elements.push(op.element);
      break;
    }
    case 'elementUpdate': {
      const s = slides[op.slide];
      if (!s || !Array.isArray(s.elements)) throw new Error('Slide index out of bounds');
      const el = s.elements[op.element];
      if (!el) throw new Error('Element index out of bounds');
      Object.assign(el, op.patch);
      break;
    }
    case 'elementDelete': {
      const s = slides[op.slide];
      if (!s || !Array.isArray(s.elements)) throw new Error('Slide index out of bounds');
      if (op.element < 0 || op.element >= s.elements.length) throw new Error('Element index out of bounds');
      s.elements.splice(op.element, 1);
      break;
    }
    case 'elementReorder': {
      const s = slides[op.slide];
      if (!s || !Array.isArray(s.elements)) throw new Error('Slide index out of bounds');
      const arr = s.elements;
      if (op.from < 0 || op.from >= arr.length || op.to < 0 || op.to >= arr.length)
        throw new Error('Element index out of bounds');
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
      throw new Error('Unknown operation: ' + JSON.stringify(_exhaustive));
    }
  }
}

function clamp(v: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, v));
}

/**
 * Trim model for sending to webview: removes large fields (image base64, __raw subtrees),
 * keeping only geometry and text info needed for editing — significantly reduces message size.
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
          delete el.data; // preview images come from slidesHtml; no need to duplicate in model
        }
        // Shape image fill (a:blipFill) embeds the same base64; rendering comes from slidesHtml
        if (el.fill && typeof el.fill === 'object' && el.fill.type === 'image') {
          delete el.fill.data;
        }
      }
    }
  }
  if (clone.media) delete clone.media;
  if (clone.fonts) delete clone.fonts;
  return clone;
}
