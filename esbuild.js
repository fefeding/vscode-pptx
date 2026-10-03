// 构建脚本：打包扩展宿主（Node）与 Webview（浏览器）两个产物。
const esbuild = require('esbuild');
const path = require('path');

const watch = process.argv.includes('--watch');
const production = process.argv.includes('--production');

async function build() {
  /** @type {import('esbuild').BuildOptions} */
  const base = {
    bundle: true,
    sourcemap: !production,
    minify: production,
    logLevel: 'info',
    legalComments: 'none'
  };

  const extension = esbuild.context
    ? null
    : null; // placeholder for older versions

  const targets = [
    esbuild.build({
      ...base,
      entryPoints: [path.join(__dirname, 'src', 'extension.ts')],
      outfile: path.join(__dirname, 'dist', 'extension.js'),
      external: ['vscode'],
      format: 'cjs',
      platform: 'node',
      target: 'node16'
    }),
    esbuild.build({
      ...base,
      entryPoints: [path.join(__dirname, 'src', 'webview', 'main.ts')],
      outfile: path.join(__dirname, 'dist', 'webview.js'),
      format: 'esm',
      platform: 'browser',
      target: 'es2020'
    })
  ];

  await Promise.all(targets);

  if (watch) {
    // 重新进入 watch 模式
    for (const t of targets) {
      if (t && typeof t === 'object' && 'then' in t) {
        // build 返回的是 promise（未使用 context），这里简单地重新构建
      }
    }
    console.log('[esbuild] initial build complete; run `node esbuild.js --watch` with context for live reload.');
  }
}

// 为支持 watch，使用 context API
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

  // 浏览器构建里解析器仍残留对 Node 内建（fs/path）的 import（仅 CLI 辅助函数使用，
  // pptxToHtml 运行时并不调用）。在 webview 中将其桩化为空实现。
  const nodeStubPlugin = {
    name: 'node-builtin-stub',
    setup(build) {
      build.onResolve({ filter: /^(fs|path|os|crypto|stream|util|http|https|zlib|url)$/ }, (args) => ({
        path: args.path,
        namespace: 'stub-node'
      }));
      build.onLoad({ filter: /.*/, namespace: 'stub-node' }, () => ({
        contents:
          'const p = new Proxy(function(){}, { get: () => () => { throw new Error("node builtin stub"); }, apply: () => { throw new Error("node builtin stub"); } });' +
          'export default p;' +
          'export const readFileSync=()=>{}, writeFileSync=()=>{}, readdirSync=()=>[], existsSync=()=>false, statSync=()=>({}), promises={};' +
          'export const join=()=>"", resolve=()=>"", dirname=()=>"", basename=()=>"";'
      }));
    }
  };

  const webCtx = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'webview', 'main.ts')],
    outfile: path.join(__dirname, 'dist', 'webview.js'),
    format: 'iife',
    platform: 'browser',
    target: 'es2020',
    // 强制使用解析器自带的浏览器构建（已内联 jszip/tinycolor2，且不依赖 Node 的 fs）
    alias: {
      '@fefeding/ppt-parser': path.resolve(__dirname, '..', 'pptx-parser', 'dist', 'ppt-parser.browser.js')
    },
    plugins: [nodeStubPlugin]
  });

  if (watch) {
    await Promise.all([extCtx.watch(), webCtx.watch()]);
    console.log('[esbuild] watching for changes…');
  } else {
    await Promise.all([extCtx.rebuild(), webCtx.rebuild()]);
    await extCtx.dispose();
    await webCtx.dispose();
    console.log('[esbuild] build complete.');
  }
}

buildWithContext().catch((e) => {
  console.error(e);
  process.exit(1);
});
