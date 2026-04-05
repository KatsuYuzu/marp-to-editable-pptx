# Marp to Editable PPTX

**[Install from VS Code Marketplace →](https://marketplace.visualstudio.com/items?itemName=KatsuYuzu.marp-to-editable-pptx)**

A VS Code extension that exports [Marp](https://marp.app/) Markdown presentations to editable PowerPoint (.pptx) files.

Each text box, image, and shape is individually placed — not embedded as a flat image — so you can freely edit the slide content in PowerPoint or LibreOffice.

> **LibreOffice is not required.**  
> Marp for VS Code's built-in editable PPTX feature depends on LibreOffice (`soffice --headless`) and is marked experimental.  
> This extension uses a browser-DOM extraction approach instead, so it works without LibreOffice installed.

## Usage

1. Open a Marp Markdown file (`.md`) in VS Code
2. Press `F1` and run **Marp: Export to Editable PPTX**
3. Choose a save location in the dialog
4. The editable `.pptx` file is generated

## Requirements

- A Chromium-based browser (Google Chrome or Microsoft Edge) must be installed

## Visual Quality

Each image shows **HTML (Marp) on the left** and **exported PPTX on the right**.  
All 60 slides from [`src/native-pptx/test-fixtures/pptx-export.md`](src/native-pptx/test-fixtures/pptx-export.md) — auto-updated by CI.

<!-- Screenshot comparison table — auto-updated by the Update Screenshots workflow -->

<details open>
<summary>All slide comparisons (60 slides)</summary>

<table>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-001.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-002.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-003.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-004.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-005.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-006.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-007.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-008.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-009.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-010.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-011.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-012.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-013.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-014.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-015.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-016.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-017.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-018.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-019.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-020.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-021.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-022.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-023.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-024.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-025.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-026.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-027.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-028.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-029.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-030.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-031.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-032.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-033.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-034.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-035.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-036.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-037.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-038.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-039.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-040.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-041.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-042.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-043.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-044.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-045.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-046.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-047.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-048.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-049.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-050.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-051.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-052.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-053.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-054.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-055.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-056.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-057.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-058.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-059.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-060.png"></td>
</tr>
</table>

</details>

## How it works

1. Converts the Markdown to HTML using [@marp-team/marp-cli](https://github.com/marp-team/marp-cli)
2. Launches a headless browser to render each slide and extract precise layout information (position, font, color, images, background)
3. Builds an editable `.pptx` where each element is individually placed as a native PowerPoint shape

See [`src/native-pptx/README.md`](src/native-pptx/README.md) for architecture details, ADR log, and the visual diff improvement workflow.

## For contributors

```sh
# Install dependencies
npm install

# Build (extension + native-pptx bundle)
npm run build

# Run unit tests
npm test

# Run the visual fidelity comparison locally (Windows, requires PowerPoint)
node src/native-pptx/tools/gen-pptx.js src/native-pptx/test-fixtures/slides-ci.html dist/compare-out.pptx
node src/native-pptx/tools/compare-visuals.js src/native-pptx/test-fixtures/slides-ci.html dist/compare-out.pptx
# → report at dist/compare-slides-ci/compare-report.html
```

## License

MIT
