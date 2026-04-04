# Marp to Editable PPTX

A VS Code extension that exports [Marp](https://marp.app/) Markdown presentations to editable PowerPoint (.pptx) files.

Each text box, image, and shape is individually placed — not embedded as a flat image — so you can freely edit the slide content in PowerPoint or LibreOffice.

## Usage

1. Open a Marp Markdown file (`.md`) in VS Code
2. Press `F1` and run **Marp: Export to Editable PPTX**
3. Choose a save location in the dialog
4. The editable `.pptx` file is generated

## Requirements

- A Chromium-based browser (Google Chrome or Microsoft Edge) must be installed

## Visual Quality

The following comparisons show Marp HTML output (left) against the exported PPTX slide rendered in PowerPoint (right).

<!-- Screenshot comparison table.
     To regenerate: run the comparison tool in src/native-pptx/tools/compare-visuals.js
     against your presentation, then copy the compare-NNN.png files into docs/screenshots/.
-->

| Slide | HTML (Marp) | PPTX (exported) | Side-by-side |
|:---:|:---:|:---:|:---:|
| 1 | ![Slide 1 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-001.png) | ![Slide 1 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-001.png) | ![Compare 1](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-001.png) |
| 2 | ![Slide 2 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-002.png) | ![Slide 2 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-002.png) | ![Compare 2](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-002.png) |
| 3 | ![Slide 3 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-003.png) | ![Slide 3 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-003.png) | ![Compare 3](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-003.png) |

> Screenshots generated from [`src/native-pptx/test-fixtures/pptx-export.md`](src/native-pptx/test-fixtures/pptx-export.md).

## How it works

1. Converts the Markdown to HTML using [@marp-team/marp-cli](https://github.com/marp-team/marp-cli)
2. Launches a headless browser to render each slide and extract precise layout information (position, font, color, images, background)
3. Builds an editable `.pptx` where each element is individually placed as a native PowerPoint shape

## License

MIT
