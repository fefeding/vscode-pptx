// 构建脚本：打包扩展宿主（Node）与 Webview（浏览器）两个产物。
//
// 用法：
//   node esbuild.js              单次构建（npm run compile）
//   node esbuild.js --watch      持续构建（npm run dev / npm run watch）
//   node esbuild.js --production 生产构建（npm run vscode:prepublish）
//
// 注意：--watch 只负责「编译」，不会拉起调试会话。要调试请在 VS Code 中按 F5，
// 并选择 "Run Extension (watch)" 启动项（它会自动先跑本 watch 任务）。
const esbuild = require('esbuild');
const path = require('path');

const watch = process.argv.includes('--watch');
const production = process.argv.includes('--production');

async function buildWithContext() {
  const base = {
    bundle: true,
    sourcemap: !production,
    minify: production,
    logLevel: 'info',
    legalComments: 'none'
  };

  const extCtx = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'extension.ts')],
    outfile: path.join(__dirname, 'dist', 'extension.js'),
    external: ['vscode'],
    format: 'cjs',
    platform: 'node',
    target: 'node16'
  });

  const webCtx = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'webview', 'main.ts')],
    outfile: path.join(__dirname, 'dist', 'webview.js'),
    format: 'iife',
    platform: 'browser',
    target: 'es2020',
    // 强制使用解析器自带的浏览器构建（已内联 jszip/tinycolor2，且不依赖 Node 的 fs）
    alias: {
      '@fefeding/ppt-parser': path.resolve(__dirname, '..', 'pptx-parser', 'dist', 'ppt-parser.browser.js'),
      // 图表渲染器由解析器包直接提供，避免在本仓库内再维护一份副本
      '@fefeding/ppt-parser/chart-renderer': path.resolve(__dirname, '..', 'pptx-parser', 'examples', 'chart-lib', 'chart-renderer.js')
    }
  });

  if (watch) {
    await Promise.all([extCtx.watch(), webCtx.watch()]);
    // 该行同时充当 .vscode/tasks.json 里 background 任务的 beginsPattern
    console.log('[esbuild] watching for changes… (press Ctrl+C to stop)');
    console.log('[esbuild] 编译已持续运行。调试：按 F5 选择 "Run Extension (watch)"；');
    console.log('[esbuild] 改代码后在扩展开发宿主窗口按 Ctrl/Cmd+R 刷新即可生效。');
    return;
  }

  await Promise.all([extCtx.rebuild(), webCtx.rebuild()]);
  await extCtx.dispose();
  await webCtx.dispose();
  console.log('[esbuild] build complete.');
}

buildWithContext().catch((e) => {
  console.error(e);
  process.exit(1);
});