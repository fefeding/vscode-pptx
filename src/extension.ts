import * as vscode from 'vscode';
import { PptxCustomDocument } from './pptx-document';
import type { HostToWebview, WebviewToHost } from './protocol';

const VIEW_TYPE = 'pptx-parser.pptxEditor';

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
      if (e.webviewPanel.active) this._activeDoc = document;
    });

    const send = (msg: HostToWebview) => webviewPanel.webview.postMessage(msg);
    const pushUpdate = async (error?: string) => {
      const snap = await document.snapshot();
      send({ type: 'update', mode: snap.mode, originalBytes: snap.originalBytes, bytes: snap.bytes, model: snap.model, canUndo: snap.canUndo, canRedo: snap.canRedo, error });
    };

    const sendInit = async () => {
      const s = await document.snapshot();
      send({
        type: 'init',
        slideSize: document.slideSize,
        mode: s.mode,
        originalBytes: s.originalBytes,
        bytes: s.bytes,
        model: s.model,
        title: document.uri.path.split('/').pop() || 'presentation.pptx',
        canUndo: s.canUndo,
        canRedo: s.canRedo
      });
    };

    // Register listener before sending, so the webview never misses a message if the host throws
    webviewPanel.webview.onDidReceiveMessage(async (message: WebviewToHost) => {
      try {
        switch (message.type) {
          case 'ready':
            await sendInit();
            break;
          case 'setMode':
            document.mode = message.mode;
            await pushUpdate();
            break;
          case 'op': {
            const res = await document.applyOp(message.op);
            if (res.ok) await pushUpdate();
            else send({ type: 'info', kind: 'error', message: res.error || 'Operation failed' });
            break;
          }
          case 'undo':
            if (await document.undo()) await pushUpdate();
            break;
          case 'redo':
            if (await document.redo()) await pushUpdate();
            break;
          case 'save':
            await document.save();
            send({ type: 'saved' });
            send({ type: 'info', kind: 'info', message: 'Saved' });
            break;
          case 'saveAs':
            await this.saveAs(document);
            break;
          case 'revert':
            await document.revert();
            await pushUpdate();
            send({ type: 'info', kind: 'info', message: 'Reloaded' });
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
      send({ type: 'info', kind: 'error', message: 'Initialization failed: ' + (err?.message || String(err)) });
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

  private async saveAs(document: PptxCustomDocument): Promise<void> {
    const uri = await vscode.window.showSaveDialog({
      defaultUri: document.uri,
      saveLabel: 'Export Copy',
      filters: { 'PowerPoint': ['pptx'] }
    });
    if (!uri) return;
    await document.save(uri);
    vscode.window.showInformationMessage(`Exported copy: ${uri.fsPath}`);
  }

  public async revertActive(): Promise<void> {
    if (this._activeDoc) await this._activeDoc.revert();
  }
  public async exportActive(): Promise<void> {
    if (this._activeDoc) await this.saveAs(this._activeDoc);
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
<html lang="en">
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

  context.subscriptions.push(
    vscode.commands.registerCommand('pptx-parser.revert', () => provider.revertActive()),
    vscode.commands.registerCommand('pptx-parser.exportCopy', () => provider.exportActive())
  );
}

export function deactivate(): void {
  // no-op
}
