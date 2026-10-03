import * as vscode from 'vscode';
import { pptxToStandard, jsonToPptx } from '@fefeding/ppt-parser';
import { applyOpToModel, cloneModel, trimModel } from './pptx-ops';
import type { PptxOp } from './protocol';

const PARSE_OPTS = { mediaProcess: true, themeProcess: true } as const;

/**
 * PPTX 自定义文档：持有标准 PptxDocument 模型（保存的事实来源）。
 *
 * 注意：扩展宿主运行在 Node 环境，没有 DOM，因此**不能**在此调用 pptxToHtml
 * （其内部 getTextWidth 依赖 document）。HTML 渲染交由 Webview（浏览器环境）完成：
 * 宿主只负责解析模型（pptxToStandard）与序列化（jsonToPptx），并把序列化字节发送给 Webview。
 */
// 在较新 @types/vscode 中 CustomDocument 是一个 interface（仅需 uri + dispose），
// 因此这里用 implements 实现，而非 extends 类。
export class PptxCustomDocument implements vscode.CustomDocument {
  private readonly _onDidChange = new vscode.EventEmitter<{ contentChanged: boolean }>();
  public readonly onDidChange = this._onDidChange.event;

  private readonly _onDidRevert = new vscode.EventEmitter<void>();
  public readonly onDidRevert = this._onDidRevert.event;

  private readonly _onDidSave = new vscode.EventEmitter<void>();
  public readonly onDidSave = this._onDidSave.event;

  public readonly uri: vscode.Uri;

  private _model!: any;
  private _slideSize: { width: number; height: number } = { width: 1280, height: 720 };
  private _originalBytes!: Uint8Array; // 文件原始字节：预览模式用此忠实渲染，不做任何往返

  private _undoStack: any[] = [];
  private _redoStack: any[] = [];
  private _dirty = false;
  private _mode: 'preview' | 'edit' = 'preview';

  private constructor(uri: vscode.Uri) {
    this.uri = uri;
  }

  static async create(uri: vscode.Uri): Promise<PptxCustomDocument> {
    const doc = new PptxCustomDocument(uri);
    const bytes = await vscode.workspace.fs.readFile(uri);
    const standard = await pptxToStandard(bytes, PARSE_OPTS);
    doc._model = standard;
    doc._originalBytes = bytes;
    doc._slideSize =
      (standard && standard.slideSize) || { width: 1280, height: 720 };
    return doc;
  }

  // ---- 只读数据访问 ----
  get slideSize() {
    return this._slideSize;
  }
  get model() {
    return this._model;
  }
  get originalBytes() {
    return this._originalBytes;
  }
  get mode() {
    return this._mode;
  }
  set mode(m: 'preview' | 'edit') {
    this._mode = m;
  }
  get isDirty() {
    return this._dirty;
  }
  get canUndo() {
    return this._undoStack.length > 0;
  }
  get canRedo() {
    return this._redoStack.length > 0;
  }

  /** 生成发送给 Webview 的 payload：原始字节（预览）+ 模型序列化字节（编辑）+ 裁剪后的模型 */
  async snapshot(): Promise<{
    mode: 'preview' | 'edit';
    originalBytes: string;
    bytes: string;
    model: any;
    canUndo: boolean;
    canRedo: boolean;
  }> {
    // 预览模式用原始字节忠实渲染，不必做昂贵的模型往返（大文件上耗时且可能失败）
    const bytes = this._mode === 'edit' ? await this.serialize() : new Uint8Array(0);
    // 必须转 base64：VSCode webview 的 typed array 传输在自定义编辑器下不会还原成 Uint8Array
    return {
      mode: this._mode,
      originalBytes: Buffer.from(this._originalBytes).toString('base64'),
      bytes: Buffer.from(bytes).toString('base64'),
      model: trimModel(this._model),
      canUndo: this.canUndo,
      canRedo: this.canRedo
    };
  }

  // ---- 编辑 ----
  /** 应用一次编辑操作。失败返回错误且不影响当前模型。 */
  async applyOp(op: PptxOp): Promise<{ ok: boolean; error?: string }> {
    const candidate = cloneModel(this._model);
    try {
      applyOpToModel(candidate, op);
    } catch (e: any) {
      return { ok: false, error: e?.message || String(e) };
    }
    // 验证可序列化后再提交（在 Node 下 jsonToPptx 不需要 DOM）
    try {
      await jsonToPptx(candidate, { outputType: 'uint8array' });
    } catch (e: any) {
      return { ok: false, error: '该操作无法生成有效 PPTX：' + (e?.message || String(e)) };
    }
    this._undoStack.push(cloneModel(this._model));
    this._redoStack = [];
    this._model = candidate;
    this._markChanged(true);
    return { ok: true };
  }

  async undo(): Promise<boolean> {
    if (!this._undoStack.length) return false;
    this._redoStack.push(cloneModel(this._model));
    this._model = this._undoStack.pop();
    this._markChanged(true);
    return true;
  }

  async redo(): Promise<boolean> {
    if (!this._redoStack.length) return false;
    this._undoStack.push(cloneModel(this._model));
    this._model = this._redoStack.pop();
    this._markChanged(true);
    return true;
  }

  private _markChanged(changed: boolean) {
    this._dirty = changed;
    this._onDidChange.fire({ contentChanged: changed });
  }

  // ---- 保存 / 回退 / 备份 ----
  async serialize(): Promise<Uint8Array> {
    return (await jsonToPptx(this._model, { outputType: 'uint8array' })) as Uint8Array;
  }

  async save(targetUri?: vscode.Uri): Promise<void> {
    const dest = targetUri || this.uri;
    // 预览模式（或未改动）直接写回原始字节，避免标准模型往返带来的保真度损失
    const bytes =
      this._mode === 'edit' && this._dirty ? await this.serialize() : this._originalBytes;
    await vscode.workspace.fs.writeFile(dest, bytes);
    if (!targetUri) {
      this._originalBytes = bytes;
      this._dirty = false;
      this._onDidChange.fire({ contentChanged: false });
      this._onDidSave.fire();
    }
  }

  async revert(): Promise<void> {
    const bytes = await vscode.workspace.fs.readFile(this.uri);
    const standard = await pptxToStandard(bytes, PARSE_OPTS);
    this._model = standard;
    this._originalBytes = bytes;
    this._slideSize = (standard && standard.slideSize) || { width: 1280, height: 720 };
    this._undoStack = [];
    this._redoStack = [];
    this._dirty = false;
    this._onDidChange.fire({ contentChanged: false });
    this._onDidRevert.fire();
  }

  async backup(destination: vscode.Uri): Promise<void> {
    const bytes = await this.serialize();
    await vscode.workspace.fs.writeFile(destination, bytes);
  }

  dispose(): void {
    this._onDidChange.dispose();
    this._onDidRevert.dispose();
    this._onDidSave.dispose();
  }
}
