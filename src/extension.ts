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

  /** 当前获得焦点的文档（用于命令如「导出副本」） */
  private _activeDoc: PptxCustomDocument | undefined;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(
    uri: vscode.Uri,
    _openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<PptxCustomDocument> {
    const document = await PptxCustomDocument.create(uri);

    // 文档内容变化时通知编辑器（用于脏标记 / 保存提示）
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

    // 先注册监听再下发，避免宿主侧异常时 webview 永远收不到消息
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
            else send({ type: 'info', kind: 'error', message: res.error || '操作失败' });
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
            send({ type: 'info', kind: 'info', message: '已保存' });
            break;
          case 'saveAs':
            await this.saveAs(document);
            break;
          case 'revert':
            await document.revert();
            await pushUpdate();
            send({ type: 'info', kind: 'info', message: '已重新加载' });
            break;
        }
      } catch (err: any) {
        send({ type: 'info', kind: 'error', message: err?.message || String(err) });
      }
    });

    // 首次下发：失败只在 webview 提示，不让整个编辑器卡在空界面
    try {
      await sendInit();
    } catch (err: any) {
      send({ type: 'info', kind: 'error', message: '初始化失败：' + (err?.message || String(err)) });
    }
  }

  // ---- CustomEditorProvider 必须实现的方法 ----
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
      saveLabel: '导出副本',
      filters: { 'PowerPoint': ['pptx'] }
    });
    if (!uri) return;
    await document.save(uri);
    vscode.window.showInformationMessage(`已导出副本：${uri.fsPath}`);
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
    // 图表依赖：echarts / echarts-gl 以全局脚本形式加载（顺序敏感，须在 webview.js 之前）
    const echartsUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts.min.js')
    );
    const echartsGlUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts-gl.min.js')
    );
    // 基础样式表：提供 .block{position:absolute}、.content{display:flex} 等版式骨架，
    // 缺失会导致所有元素退化为纵向堆叠（与 examples/index.html 的渲染差距主要来自这里）
    const baseCssUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'media', 'pptxjs.css')
    );
    const nonce = getNonce();
    const csp = [
      `default-src 'none';`,
      `img-src ${webview.cspSource} data: blob:;`,
      `media-src ${webview.cspSource} data: blob:;`,
      `style-src ${webview.cspSource} 'unsafe-inline';`,
      // 'unsafe-eval' 是 echarts-gl 的硬性要求：它内部用 new Function() 解析 expr() 表达式，
      // 否则 3D 图表会抛 "Invalid expression."（CSP 拦截后被 catch 吞掉真实原因）
      `script-src 'nonce-${nonce}' 'unsafe-eval' ${webview.cspSource};`
    ].join(' ');
    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Security-Policy" content="${csp}" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>PPTX Editor</title>
<link rel="stylesheet" href="${baseCssUri}" />
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
