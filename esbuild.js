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
const fs = require('fs');

/**
 * Resolve an entry file inside @fefeding/ppt-parser.
 * Prefer a sibling `pptx-parser` repo (local dev / live-editing the parser), and fall back to the
 * published npm package in node_modules (used on CI or when no sibling repo is present). The npm
 * package already ships `dist/ppt-parser.browser.js` and `examples/chart-lib/chart-renderer.js`.
 */
function resolveParserEntry(rel) {
  const local = path.resolve(__dirname, '..', 'pptx-parser', rel);
  if (fs.existsSync(local)) return local;
  const fromNodeModules = path.resolve(__dirname, 'node_modules', '@fefeding', 'ppt-parser', rel);
  if (!fs.existsSync(fromNodeModules)) {
    throw new Error(
      `Cannot resolve @fefeding/ppt-parser entry "${rel}". ` +
      `Run "npm install" or place a sibling "pptx-parser" repo so the file is available.`
    );
  }
  return fromNodeModules;
}

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
    // 优先用同级 pptx-parser 仓库的本地构建（开发期联调），CI/无本地仓库时回退到 node_modules 中
    // 已发布的 npm 包（@fefeding/ppt-parser 自带 dist 浏览器构建与 chart-renderer）。
    alias: {
      '@fefeding/ppt-parser': resolveParserEntry('dist/ppt-parser.browser.js'),
      '@fefeding/ppt-parser/chart-renderer': resolveParserEntry('examples/chart-lib/chart-renderer.js')
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