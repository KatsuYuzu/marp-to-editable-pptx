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

Each image shows **HTML (Marp) on the left** and **exported PPTX on the right**.  
All 59 slides from [`src/native-pptx/test-fixtures/pptx-export.md`](src/native-pptx/test-fixtures/pptx-export.md) — auto-updated by CI.

<!-- Screenshot comparison table — auto-updated by the Update Screenshots workflow -->

<details open>
<summary>All slide comparisons (59 slides)</summary>

| Slide 1 | Slide 2 |
|:---:|:---:|
| ![1](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-001.png) | ![2](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-002.png) |
| ![3](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-003.png) | ![4](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-004.png) |
| ![5](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-005.png) | ![6](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-006.png) |
| ![7](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-007.png) | ![8](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-008.png) |
| ![9](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-009.png) | ![10](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-010.png) |
| ![11](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-011.png) | ![12](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-012.png) |
| ![13](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-013.png) | ![14](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-014.png) |
| ![15](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-015.png) | ![16](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-016.png) |
| ![17](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-017.png) | ![18](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-018.png) |
| ![19](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-019.png) | ![20](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-020.png) |
| ![21](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-021.png) | ![22](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-022.png) |
| ![23](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-023.png) | ![24](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-024.png) |
| ![25](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-025.png) | ![26](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-026.png) |
| ![27](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-027.png) | ![28](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-028.png) |
| ![29](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-029.png) | ![30](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-030.png) |
| ![31](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-031.png) | ![32](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-032.png) |
| ![33](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-033.png) | ![34](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-034.png) |
| ![35](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-035.png) | ![36](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-036.png) |
| ![37](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-037.png) | ![38](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-038.png) |
| ![39](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-039.png) | ![40](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-040.png) |
| ![41](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-041.png) | ![42](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-042.png) |
| ![43](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-043.png) | ![44](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-044.png) |
| ![45](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-045.png) | ![46](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-046.png) |
| ![47](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-047.png) | ![48](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-048.png) |
| ![49](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-049.png) | ![50](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-050.png) |
| ![51](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-051.png) | ![52](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-052.png) |
| ![53](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-053.png) | ![54](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-054.png) |
| ![55](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-055.png) | ![56](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-056.png) |
| ![57](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-057.png) | ![58](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-058.png) |
| ![59](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-059.png) | |

</details>

## How it works

1. Converts the Markdown to HTML using [@marp-team/marp-cli](https://github.com/marp-team/marp-cli)
2. Launches a headless browser to render each slide and extract precise layout information (position, font, color, images, background)
3. Builds an editable `.pptx` where each element is individually placed as a native PowerPoint shape

## License

MIT
