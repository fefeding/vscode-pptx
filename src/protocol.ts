// Message protocol between extension host (Node) and webview (browser).

/** A single edit operation. */
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

/** Webview → Host */
export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'op'; op: PptxOp }
  | { type: 'undo' }
  | { type: 'redo' }
  | { type: 'save' }
  | { type: 'saveAs' }
  | { type: 'revert' }
  | { type: 'setMode'; mode: 'preview' | 'edit' };

/** Host → Webview */
export type HostToWebview =
  | {
      type: 'init';
      slideSize: { width: number; height: number };
      mode: 'preview' | 'edit'; // initial mode (default: preview)
      // Transmitted as base64: VSCode webview typed array transport ($$vscode_array_buffer_reference$$)
      // does not restore to Uint8Array in custom editor scenarios; passing binary directly causes parse failure.
      originalBytes: string; // original file bytes (base64) — used by preview mode for faithful rendering via pptxToHtml
      bytes: string; // model-serialized PPTX bytes (base64) — used by edit mode rendering (standard model round-trip)
      model: any; // trimmed standard PptxDocument (large fields removed)
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
