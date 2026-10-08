# PPTX Editor for VSCode

A VSCode extension built on [`@fefeding/ppt-parser`](https://github.com/fefeding/pptx-parser) that lets you **open and edit `.pptx` files visually** — directly inside the editor.

## Features

- **Default `.pptx` handler**: Double-click any `.pptx` in the Explorer to open it with the built-in visual editor (no extra software required).
- **Slide preview**: Renders every slide at its original aspect ratio via `pptxToHtml` (most faithful). Left sidebar thumbnail navigation with zoom-to-fit and grid view.
- **Visual editing**:
  - Click to select elements on the canvas; drag to move; drag handles to resize; drag the rotation grip to rotate.
  - Properties panel: position / size / rotation, text (content / font size / color / alignment / bold / italic / underline), shapes (fill / border), image replacement.
  - Slide operations: add blank slide, duplicate, delete, reorder up/down, set background color, lock, hide, speaker notes.
  - Insert elements: text box, rectangle, image (from local file).
  - Alignment / distribution / grouping / z-order / find-replace via the editor core actions.
- **Presentation mode**: Press `F5` or click "Present" for fullscreen slideshow. Navigate with arrow keys, space, or click.
- **Save & export**: `Ctrl/Cmd+S` to save; "Export Copy…" to save a copy; toolbar "Discard Changes & Reload".
- **Undo / Redo**: `Ctrl/Cmd+Z`, `Ctrl/Cmd+Y`.

## Development

```bash
npm install            # installs dependencies, including @fefeding/ppt-parser from npm
npm run compile        # esbuild bundles dist/extension.js and dist/webview.js
npm run watch          # watch mode (alias: npm run dev)
npm run package        # build the production bundle and package a .vsix (vsce package)
npm run publish        # publish to the VS Code Marketplace (vsce publish)
```

Press `F5` in VSCode to launch an Extension Development Host, then open any `.pptx` file.
The default zoom for both preview and edit canvases is "fit to window"; use the status-bar
zoom controls (or `Ctrl/Cmd` + wheel on the canvas) for manual zoom.

## Notes & Limitations

- Editing renders straight from the editor document model provided by `@fefeding/ppt-parser`
  (its headless editor core: `createStore` / `docFromPptx` / `createActions` / `renderSlideInto`).
  On save the document is serialized back to PPTX; **slide masters / layouts** are rewritten to a
  single blank layout. **Content elements** — text boxes, shapes, images, charts, tables — are
  editable and preserved, and because editing renders from the model (not a byte round-trip) it
  never loses fidelity such as picture or pattern fills.
- The preview mode uses `pptxToHtml` for the most faithful rendering of the original file.
- Charts are rendered with ECharts (the same renderer the parser's preview uses), including 3D
  charts where supported by the source.
- Complex objects such as SmartArt, OLE, and equations are retained as raw OOXML (`__raw`) but are
  not visually editable.
