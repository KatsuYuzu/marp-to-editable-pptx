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

The following comparisons show Marp HTML output (left) against the exported PPTX slide rendered by LibreOffice (right).
All 59 slides from [`src/native-pptx/test-fixtures/pptx-export.md`](src/native-pptx/test-fixtures/pptx-export.md) are shown.

<!-- Screenshot comparison table — auto-updated by the Update Screenshots workflow -->

<details open>
<summary>All slide comparisons (59 slides)</summary>

| Slide | HTML (Marp) | PPTX (exported) | Side-by-side |
|:---:|:---:|:---:|:---:|
| 1 | ![1 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-001.png) | ![1 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-001.png) | ![Compare 1](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-001.png) |
| 2 | ![2 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-002.png) | ![2 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-002.png) | ![Compare 2](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-002.png) |
| 3 | ![3 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-003.png) | ![3 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-003.png) | ![Compare 3](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-003.png) |
| 4 | ![4 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-004.png) | ![4 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-004.png) | ![Compare 4](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-004.png) |
| 5 | ![5 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-005.png) | ![5 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-005.png) | ![Compare 5](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-005.png) |
| 6 | ![6 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-006.png) | ![6 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-006.png) | ![Compare 6](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-006.png) |
| 7 | ![7 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-007.png) | ![7 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-007.png) | ![Compare 7](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-007.png) |
| 8 | ![8 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-008.png) | ![8 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-008.png) | ![Compare 8](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-008.png) |
| 9 | ![9 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-009.png) | ![9 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-009.png) | ![Compare 9](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-009.png) |
| 10 | ![10 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-010.png) | ![10 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-010.png) | ![Compare 10](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-010.png) |
| 11 | ![11 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-011.png) | ![11 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-011.png) | ![Compare 11](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-011.png) |
| 12 | ![12 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-012.png) | ![12 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-012.png) | ![Compare 12](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-012.png) |
| 13 | ![13 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-013.png) | ![13 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-013.png) | ![Compare 13](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-013.png) |
| 14 | ![14 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-014.png) | ![14 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-014.png) | ![Compare 14](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-014.png) |
| 15 | ![15 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-015.png) | ![15 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-015.png) | ![Compare 15](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-015.png) |
| 16 | ![16 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-016.png) | ![16 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-016.png) | ![Compare 16](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-016.png) |
| 17 | ![17 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-017.png) | ![17 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-017.png) | ![Compare 17](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-017.png) |
| 18 | ![18 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-018.png) | ![18 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-018.png) | ![Compare 18](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-018.png) |
| 19 | ![19 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-019.png) | ![19 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-019.png) | ![Compare 19](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-019.png) |
| 20 | ![20 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-020.png) | ![20 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-020.png) | ![Compare 20](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-020.png) |
| 21 | ![21 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-021.png) | ![21 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-021.png) | ![Compare 21](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-021.png) |
| 22 | ![22 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-022.png) | ![22 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-022.png) | ![Compare 22](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-022.png) |
| 23 | ![23 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-023.png) | ![23 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-023.png) | ![Compare 23](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-023.png) |
| 24 | ![24 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-024.png) | ![24 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-024.png) | ![Compare 24](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-024.png) |
| 25 | ![25 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-025.png) | ![25 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-025.png) | ![Compare 25](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-025.png) |
| 26 | ![26 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-026.png) | ![26 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-026.png) | ![Compare 26](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-026.png) |
| 27 | ![27 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-027.png) | ![27 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-027.png) | ![Compare 27](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-027.png) |
| 28 | ![28 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-028.png) | ![28 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-028.png) | ![Compare 28](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-028.png) |
| 29 | ![29 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-029.png) | ![29 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-029.png) | ![Compare 29](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-029.png) |
| 30 | ![30 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-030.png) | ![30 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-030.png) | ![Compare 30](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-030.png) |
| 31 | ![31 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-031.png) | ![31 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-031.png) | ![Compare 31](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-031.png) |
| 32 | ![32 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-032.png) | ![32 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-032.png) | ![Compare 32](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-032.png) |
| 33 | ![33 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-033.png) | ![33 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-033.png) | ![Compare 33](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-033.png) |
| 34 | ![34 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-034.png) | ![34 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-034.png) | ![Compare 34](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-034.png) |
| 35 | ![35 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-035.png) | ![35 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-035.png) | ![Compare 35](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-035.png) |
| 36 | ![36 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-036.png) | ![36 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-036.png) | ![Compare 36](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-036.png) |
| 37 | ![37 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-037.png) | ![37 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-037.png) | ![Compare 37](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-037.png) |
| 38 | ![38 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-038.png) | ![38 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-038.png) | ![Compare 38](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-038.png) |
| 39 | ![39 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-039.png) | ![39 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-039.png) | ![Compare 39](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-039.png) |
| 40 | ![40 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-040.png) | ![40 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-040.png) | ![Compare 40](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-040.png) |
| 41 | ![41 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-041.png) | ![41 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-041.png) | ![Compare 41](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-041.png) |
| 42 | ![42 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-042.png) | ![42 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-042.png) | ![Compare 42](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-042.png) |
| 43 | ![43 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-043.png) | ![43 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-043.png) | ![Compare 43](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-043.png) |
| 44 | ![44 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-044.png) | ![44 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-044.png) | ![Compare 44](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-044.png) |
| 45 | ![45 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-045.png) | ![45 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-045.png) | ![Compare 45](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-045.png) |
| 46 | ![46 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-046.png) | ![46 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-046.png) | ![Compare 46](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-046.png) |
| 47 | ![47 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-047.png) | ![47 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-047.png) | ![Compare 47](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-047.png) |
| 48 | ![48 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-048.png) | ![48 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-048.png) | ![Compare 48](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-048.png) |
| 49 | ![49 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-049.png) | ![49 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-049.png) | ![Compare 49](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-049.png) |
| 50 | ![50 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-050.png) | ![50 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-050.png) | ![Compare 50](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-050.png) |
| 51 | ![51 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-051.png) | ![51 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-051.png) | ![Compare 51](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-051.png) |
| 52 | ![52 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-052.png) | ![52 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-052.png) | ![Compare 52](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-052.png) |
| 53 | ![53 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-053.png) | ![53 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-053.png) | ![Compare 53](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-053.png) |
| 54 | ![54 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-054.png) | ![54 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-054.png) | ![Compare 54](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-054.png) |
| 55 | ![55 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-055.png) | ![55 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-055.png) | ![Compare 55](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-055.png) |
| 56 | ![56 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-056.png) | ![56 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-056.png) | ![Compare 56](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-056.png) |
| 57 | ![57 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-057.png) | ![57 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-057.png) | ![Compare 57](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-057.png) |
| 58 | ![58 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-058.png) | ![58 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-058.png) | ![Compare 58](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-058.png) |
| 59 | ![59 HTML](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/html-slide-059.png) | ![59 PPTX](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/pptx-slide-059.png) | ![Compare 59](https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-059.png) |

</details>

## How it works

1. Converts the Markdown to HTML using [@marp-team/marp-cli](https://github.com/marp-team/marp-cli)
2. Launches a headless browser to render each slide and extract precise layout information (position, font, color, images, background)
3. Builds an editable `.pptx` where each element is individually placed as a native PowerPoint shape

## License

MIT
