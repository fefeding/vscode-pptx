import * as vscode from 'vscode';
import { pptxToStandard, jsonToPptx } from '@fefeding/ppt-parser';
import { applyOpToModel, cloneModel, trimModel } from './pptx-ops';
import type { PptxOp } from './protocol';

const PARSE_OPTS = { mediaProcess: true, themeProcess: true } as const;

/**
 * PPTX custom document: holds the standard PptxDocument model (source of truth for saving).
 *
 * Note: the extension host runs in Node without a DOM, so pptxToHtml **must not** be called here
 * (its internal getTextWidth depends on document). HTML rendering is delegated to the webview
 * (browser environment): the host only parses the model (pptxToStandard) and serializes (jsonToPptx),
 * then sends the serialized bytes to the webview.
 */
// In newer @types/vscode, CustomDocument is an interface (only needs uri + dispose),
// so we use implements rather than extends.
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
  private _originalBytes!: Uint8Array; // original file bytes: preview mode renders these faithfully, no round-trip

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

  // ---- Read-only data access ----
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

  /** Generate the payload sent to the webview: original bytes (preview) + model-serialized bytes (edit) + trimmed model. */
  async snapshot(): Promise<{
    mode: 'preview' | 'edit';
    originalBytes: string;
    bytes: string;
    model: any;
    canUndo: boolean;
    canRedo: boolean;
  }> {
    // Preview mode renders original bytes faithfully; skip expensive model round-trip (slow and may fail on large files)
    const bytes = this._mode === 'edit' ? await this.serialize() : new Uint8Array(0);
    // Must convert to base64: VSCode webview typed array transport does not restore to Uint8Array in custom editors
    return {
      mode: this._mode,
      originalBytes: Buffer.from(this._originalBytes).toString('base64'),
      bytes: Buffer.from(bytes).toString('base64'),
      model: trimModel(this._model),
      canUndo: this.canUndo,
      canRedo: this.canRedo
    };
  }

  // ---- Editing ----
  /** Apply an edit operation. Returns error on failure without affecting the current model. */
  async applyOp(op: PptxOp): Promise<{ ok: boolean; error?: string }> {
    const candidate = cloneModel(this._model);
    try {
      applyOpToModel(candidate, op);
    } catch (e: any) {
      return { ok: false, error: e?.message || String(e) };
    }
    // Validate serializability before committing (jsonToPptx does not need DOM in Node)
    try {
      await jsonToPptx(candidate, { outputType: 'uint8array' });
    } catch (e: any) {
      return { ok: false, error: 'This operation cannot produce a valid PPTX: ' + (e?.message || String(e)) };
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

  // ---- Save / Revert / Backup ----
  async serialize(): Promise<Uint8Array> {
    return (await jsonToPptx(this._model, { outputType: 'uint8array' })) as Uint8Array;
  }

  async save(targetUri?: vscode.Uri): Promise<void> {
    const dest = targetUri || this.uri;
    // Preview mode (or unmodified) writes back original bytes to avoid fidelity loss from model round-trip
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
