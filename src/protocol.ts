// 扩展宿主（Node）与 Webview（浏览器）之间的消息协议

/** 一次编辑操作 */
export type PptxOp =
  | { kind: 'slideAdd'; after?: number; slide?: any }
  | { kind: 'slideDelete'; index: number }
  | { kind: 'slideDuplicate'; index: number }
  | { kind: 'slideMove'; from: number; to: number }
  | { kind: 'slideUpdate'; index: number; patch: any }
  | { kind: 'elementAdd'; slide: number; element: any }
  | { kind: 'elementUpdate'; slide: number; element: number; patch: any }
  | { kind: 'elementDelete'; slide: number; element: number }
  | { kind: 'elementReorder'; slide: number; from: number; to: number }
  | { kind: 'metadataUpdate'; patch: any };

/** Webview → 宿主 */
export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'op'; op: PptxOp }
  | { type: 'undo' }
  | { type: 'redo' }
  | { type: 'save' }
  | { type: 'saveAs' }
  | { type: 'revert' }
  | { type: 'setMode'; mode: 'preview' | 'edit' };

/** 宿主 → Webview */
export type HostToWebview =
  | {
      type: 'init';
      slideSize: { width: number; height: number };
      mode: 'preview' | 'edit'; // 初始模式（默认 preview）
      // 以 base64 传输：VSCode webview 的 typed array 传输（$$vscode_array_buffer_reference$$）在
      // 自定义编辑器场景下不会还原成 Uint8Array，直接传二进制会导致解析失败
      originalBytes: string; // 文件原始字节（base64）—— 预览模式用 pptxToHtml 忠实渲染
      bytes: string; // 模型序列化后的 PPTX 字节（base64）—— 编辑模式渲染用（标准模型往返）
      model: any; // 已裁剪的标准 PptxDocument（去除大体积字段）
      title: string;
      canUndo: boolean;
      canRedo: boolean;
    }
  | {
      type: 'update';
      mode: 'preview' | 'edit';
      originalBytes: string;
      bytes: string;
      model: any;
      canUndo: boolean;
      canRedo: boolean;
      error?: string;
    }
  | { type: 'saved' }
  | { type: 'info'; message: string; kind?: 'info' | 'warn' | 'error' };
