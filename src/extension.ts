import * as vscode from 'vscode';
import { PDFDocument } from 'pdf-lib';
import { PptxCustomDocument } from './pptx-document';
import type { EditorMode, HostToWebview, WebviewToHost } from './protocol';

const VIEW_TYPE = 'pptx-parser.pptxEditor';

/**
 * Localized host-side string. `vscode.l10n.t` echoes the key back when no bundle is installed
 * (or the locale is untranslated), so the English fallback is kept as a second argument.
 */
function loc(key: string, fallback: string, ...args: string[]): string {
  const value = vscode.l10n.t(key, ...args);
  return !value || value === key ? fallback.replace('{0}', args[0] ?? '') : value;
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

class PptxEditorProvider implements vscode.CustomEditorProvider<PptxCustomDocument> {
  public static readonly viewType = VIEW_TYPE;

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
    vscode.CustomDocumentEditEvent<PptxCustomDocument> | vscode.CustomDocumentContentChangeEvent<PptxCustomDocument>
  >();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  /** The currently focused document (used by commands like "Export Copy"). */
  private _activeDoc: PptxCustomDocument | undefined;

  /** The webview panel of the currently focused document (so title-bar commands can message it). */
  private _activePanel: vscode.WebviewPanel | undefined;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(
    uri: vscode.Uri,
    _openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<PptxCustomDocument> {
    const document = await PptxCustomDocument.create(uri);

    // Notify the editor when document content changes (for dirty indicator / save prompt)
    document.onDidChange(() => this._onDidChangeCustomDocument.fire({ document }));
    document.onDidRevert(() => this._onDidChangeCustomDocument.fire({ document }));
    document.onDidSave(() => this._onDidChangeCustomDocument.fire({ document }));
    return document;
  }

  async resolveCustomEditor(
    document: PptxCustomDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    webviewPanel.webview.options = {
      enableScripts: true,
      enableCommandUris: false,
      localResourceRoots: [this.context.extensionUri]
    };
    webviewPanel.webview.html = this.getWebviewHtml(webviewPanel.webview);

    webviewPanel.onDidChangeViewState((e) => {
      if (e.webviewPanel.active) {
        this._activeDoc = document;
        this._activePanel = webviewPanel;
        this.syncContext();
      }
    });

    const send = (msg: HostToWebview) => webviewPanel.webview.postMessage(msg);
    const sendInit = async () => {
      const s = document.snapshot();
      send({
        type: 'init',
        mode: s.mode,
        originalBytes: s.originalBytes,
        bytes: s.bytes,
        title: document.uri.path.split('/').pop() || 'presentation.pptx',
        dirty: document.isDirty,
        locale: vscode.env.language
      });
    };

    // Surface the host's real dirty state so the webview's dot matches VS Code's own save/revert.
    // The webview no longer implements saving (VS Code handles Ctrl+S / File › Save), so the
    // authoritative dirty flag lives here and is pushed down on every change.
    document.onDidSave(() => send({ type: 'saved' }));
    document.onDidRevert(() => {
      void sendInit();
      send({ type: 'reverted' });
    });

    // Register listener before sending, so the webview never misses a message if the host throws
    webviewPanel.webview.onDidReceiveMessage(async (message: WebviewToHost) => {
      try {
        switch (message.type) {
          case 'ready':
            await sendInit();
            break;
          case 'setMode':
            document.mode = message.mode;
            // Keep the `pptx:editing` context key in sync with the active document's mode.
            if (this._activeDoc === document) this.syncContext();
            break;
          case 'sync':
            document.setBytes(message.bytes);
            document.setDirty(message.dirty);
            break;
          case 'save':
            // Saving is owned by VS Code; this path is only kept for completeness. The actual
            // dirty-state notification is pushed via the document's onDidSave listener above.
            await document.save();
            break;
          case 'saveAs':
            await this.saveAs(document, message.bytes);
            break;
          case 'exportImage':
            await this.saveImageFile(document, message);
            break;
          case 'exportImages':
            await this.saveImageFiles(document, message);
            break;
          case 'exportPdf':
            await this.savePdf(document, message);
            break;
          case 'revert':
            await document.revert();
            await sendInit();
            send({ type: 'info', kind: 'info', message: loc('extension.reloaded', 'Reloaded') });
            break;
          case 'error':
            send({ type: 'info', kind: 'error', message: message.message });
            break;
        }
      } catch (err: any) {
        send({ type: 'info', kind: 'error', message: err?.message || String(err) });
      }
    });

    // First send: failures only show a toast in the webview; don't let the entire editor stall on an empty view
    try {
      await sendInit();
    } catch (err: any) {
      send({ type: 'info', kind: 'error', message: loc('extension.initFailed', 'Initialization failed: {0}', String(err?.message || err)) });
    }
  }

  // ---- CustomEditorProvider required methods ----
  async saveCustomDocument(document: PptxCustomDocument, _token: vscode.CancellationToken): Promise<void> {
    await document.save();
  }
  async saveCustomDocumentAs(
    document: PptxCustomDocument,
    target: vscode.Uri,
    _token: vscode.CancellationToken
  ): Promise<void> {
    await document.save(target);
  }
  async revertCustomDocument(document: PptxCustomDocument, _token: vscode.CancellationToken): Promise<void> {
    await document.revert();
  }
  async backupCustomDocument(
    document: PptxCustomDocument,
    context: vscode.CustomDocumentBackupContext,
    _token: vscode.CancellationToken
  ): Promise<vscode.CustomDocumentBackup> {
    await document.backup(context.destination);
    return {
      id: context.destination.toString(),
      delete: () => {
        vscode.workspace.fs.delete(context.destination).then(
          () => undefined,
          () => undefined
        );
      }
    };
  }

  private async saveAs(document: PptxCustomDocument, bytes?: string | null): Promise<void> {
    const uri = await vscode.window.showSaveDialog({
      defaultUri: document.uri,
      saveLabel: loc('extension.exportCopy', 'Export Copy'),
      filters: { 'PowerPoint': ['pptx'] }
    });
    if (!uri) return;
    // Adopt the webview's latest bytes when provided, then export (saving to another
    // targetUri intentionally leaves the dirty flag untouched)
    if (bytes) document.setBytes(bytes);
    await document.save(uri);
    vscode.window.showInformationMessage(loc('extension.exported', 'Exported copy: {0}', uri.fsPath));
  }

  // ---- Export targets (PNG / PDF): the webview renders slides offscreen, the host writes files ----

  private async saveImageFile(
    document: PptxCustomDocument,
    msg: { bytes: string; suggestedName: string }
  ): Promise<void> {
    const uri = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.joinPath(document.uri, '..', msg.suggestedName || 'slide.png'),
      filters: { 'PNG': ['png'] }
    });
    if (!uri) return;
    await vscode.workspace.fs.writeFile(uri, Buffer.from(msg.bytes, 'base64'));
    vscode.window.showInformationMessage(loc('extension.imageSaved', 'Image saved: {0}', uri.fsPath));
  }

  private async saveImageFiles(
    document: PptxCustomDocument,
    msg: { images: { name: string; bytes: string }[] }
  ): Promise<void> {
    const folders = await vscode.window.showOpenDialog({
      canSelectFiles: false,
      canSelectFolders: true,
      canSelectMany: false,
      defaultUri: vscode.Uri.joinPath(document.uri, '..'),
      openLabel: loc('extension.chooseFolder', 'Choose Folder')
    });
    if (!folders?.length) return;
    const dir = folders[0];
    for (const img of msg.images) {
      await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(dir, img.name), Buffer.from(img.bytes, 'base64'));
    }
    vscode.window.showInformationMessage(
      loc('extension.imagesSaved', 'Saved {0} images to {1}', String(msg.images.length), dir.fsPath)
    );
  }

  private async savePdf(
    document: PptxCustomDocument,
    msg: { images: { bytes: string }[]; widthPx: number; heightPx: number; suggestedName: string }
  ): Promise<void> {
    if (!msg.images?.length) return;
    const pdf = await PDFDocument.create();
    // px → pt (1px = 0.75pt, the same conversion the parser uses elsewhere)
    const w = (msg.widthPx || 1280) * 0.75;
    const h = (msg.heightPx || 720) * 0.75;
    for (const img of msg.images) {
      const embedded = await pdf.embedPng(Buffer.from(img.bytes, 'base64'));
      const page = pdf.addPage([w, h]);
      page.drawImage(embedded, { x: 0, y: 0, width: w, height: h });
    }
    const uri = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.joinPath(document.uri, '..', msg.suggestedName || 'slides.pdf'),
      filters: { 'PDF': ['pdf'] }
    });
    if (!uri) return;
    await vscode.workspace.fs.writeFile(uri, Buffer.from(await pdf.save()));
    vscode.window.showInformationMessage(loc('extension.pdfSaved', 'PDF saved: {0}', uri.fsPath));
  }

  public async revertActive(): Promise<void> {
    if (this._activeDoc) await this._activeDoc.revert();
  }
  public async exportActive(): Promise<void> {
    if (this._activeDoc) await this.saveAs(this._activeDoc);
  }

  /** 刷新编辑器标题栏切换所需的上下文键（pptx:editing）。 */
  private syncContext(): void {
    void vscode.commands.executeCommand('setContext', 'pptx:editing', this._activeDoc?.mode === 'edit');
  }

  /** 标题栏「预览 / 编辑」命令：把模式推给当前活动的 webview。 */
  public async setActiveMode(mode: EditorMode): Promise<void> {
    const doc = this._activeDoc;
    if (!doc) return;
    doc.mode = mode;
    this.syncContext();
    await this._activePanel?.webview.postMessage({ type: 'setMode', mode } as HostToWebview);
  }

  /** 标题栏切换按钮：在预览 / 编辑之间翻转当前活动文档的模式。 */
  public async toggleActiveMode(): Promise<void> {
    const doc = this._activeDoc;
    if (!doc) return;
    await this.setActiveMode(doc.mode === 'edit' ? 'preview' : 'edit');
  }

  private getWebviewHtml(webview: vscode.Webview): string {
    const scriptUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview.js')
    );
    // Chart dependencies: echarts / echarts-gl loaded as global scripts (order-sensitive, must precede webview.js)
    const echartsUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts.min.js')
    );
    const echartsGlUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts-gl.min.js')
    );
    // Base layout skeleton (.block/.content etc.) now comes from the parser's styles.global,
    // injected by the webview via ensureGlobalStyles().
    const nonce = getNonce();
    const csp = [
      `default-src 'none';`,
      `img-src ${webview.cspSource} data: blob:;`,
      `media-src ${webview.cspSource} data: blob:;`,
      `style-src ${webview.cspSource} 'unsafe-inline';`,
      // 'unsafe-eval' is required by echarts-gl: it uses new Function() to parse expr() expressions,
      // otherwise 3D charts throw "Invalid expression." (CSP blocks it and the real cause is swallowed by catch)
      `script-src 'nonce-${nonce}' 'unsafe-eval' ${webview.cspSource};`
    ].join(' ');
    return `<!DOCTYPE html>
<html lang="${vscode.env.language}">
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Security-Policy" content="${csp}" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>PPTX Editor</title>
</head>
<body>
<div id="app"></div>
<script nonce="${nonce}" src="${echartsUri}"></script>
<script nonce="${nonce}" src="${echartsGlUri}"></script>
<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}

export function activate(context: vscode.ExtensionContext): void {
  const provider = new PptxEditorProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false
    })
  );

  // Seed `pptx:editing` with a concrete boolean so the context key is never undefined.
  void vscode.commands.executeCommand('setContext', 'pptx:editing', false);

  context.subscriptions.push(
    vscode.commands.registerCommand('pptx-parser.revert', () => provider.revertActive()),
    vscode.commands.registerCommand('pptx-parser.exportCopy', () => provider.exportActive()),
    vscode.commands.registerCommand('pptx-parser.viewMode', () => provider.setActiveMode('preview')),
    vscode.commands.registerCommand('pptx-parser.editMode', () => provider.setActiveMode('edit')),
    vscode.commands.registerCommand('pptx-parser.toggleMode', () => provider.toggleActiveMode())
  );
}

export function deactivate(): void {
  // no-op
}
