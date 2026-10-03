# PPTX Editor for VSCode

A VSCode extension built on [`@fefeding/ppt-parser`](https://github.com/fefeding/pptx-parser) that lets you **open and edit `.pptx` files visually** — directly inside the editor.

## Features

- **Default `.pptx` handler**: Double-click any `.pptx` in the Explorer to open it with the built-in visual editor (no extra software required).
- **Slide preview**: Renders every slide at its original aspect ratio. Left sidebar thumbnail navigation with zoom-to-fit and grid view.
- **Visual editing**:
  - Click to select elements on the canvas; drag to move; drag handles to resize.
  - Properties panel: position / size / rotation, text (content / font size / color / alignment / bold / italic / underline), shapes (fill / border), image replacement.
  - Slide operations: add blank slide, duplicate, delete, reorder up/down, set background color, hide, speaker notes.
  - Insert elements: text box, rectangle, image (from local file).
- **Presentation mode**: Press `F5` or click "Present" for fullscreen slideshow. Navigate with arrow keys or click.
- **Save & export**: `Ctrl/Cmd+S` to save; "Export" to save a copy; toolbar "Discard Changes & Reload".
- **Undo / Redo**: `Ctrl/Cmd+Z`, `Ctrl/Cmd+Y`.

## Development

```bash
npm install            # installs pptx-parser (file:../pptx-parser) and dependencies
npm run compile        # esbuild bundles dist/extension.js and dist/webview.js
npm run watch          # watch mode
```

Press `F5` in VSCode to launch an Extension Development Host, then open any `.pptx` file.

## Notes & Limitations

- Editing uses a bidirectional round-trip between a standard JSON model (`PptxDocument`) and PPTX (provided by pptx-parser).
  On save, the presentation is regenerated; **slide masters / layouts** are rewritten to a single blank layout.
  **Content elements** — text boxes, shapes, images, charts, tables — are editable and preserved.
- Complex objects such as SmartArt, OLE, and equations are retained as raw OOXML (`__raw`) but are not visually editable.
