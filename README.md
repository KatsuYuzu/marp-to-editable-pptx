# Marp to Editable PPTX

[![Install](https://img.shields.io/badge/VS%20Code-Install-007ACC?logo=visual-studio-code)](https://marketplace.visualstudio.com/items?itemName=KatsuYuzu.marp-to-editable-pptx)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

> Export your [Marp](https://marp.app/) Markdown presentations to **truly editable PowerPoint files** — no LibreOffice, no extra software.

Each text box, image, and shape is placed as an individual native PowerPoint object, so you can freely move, resize, and restyle content in PowerPoint after export.

---

## Why this extension?

The official Marp toolchain offers two PPTX export modes, both with limitations:

| | Marp: PPTX (screenshot) | Marp: PPTX (LibreOffice) | **This extension** |
|---|:---:|:---:|:---:|
| Text/shapes are editable | ❌ flat image | ⚠️ messy PDF-converted objects | ✅ clean native objects |
| Requires LibreOffice | — | ❌ must install | ✅ not needed |
| Works in enterprise environments | ✅ | ❌ | ✅ |
| Layout faithfulness | ✅ | ⚠️ | ⚠️ |

The LibreOffice-based export works by converting the slide to a PDF and then importing it into PowerPoint via LibreOffice. This results in cluttered, hard-to-edit objects. More importantly, **many enterprise users cannot install LibreOffice** due to IT policies.

This extension uses a different approach: it reads the browser-rendered DOM directly, extracting the exact position, font, color, and content of every element, and builds a native PPTX from scratch.

> **Layout faithfulness note:** Text may look slightly different or wrap at a different point in PowerPoint than in the browser. This happens because browsers and PowerPoint use separate text rendering engines with different character spacing and line-break calculations — even with the same font. The effect is strongest when a Marp theme uses **web fonts** (e.g. from Google Fonts): PowerPoint substitutes a system font, changing character widths enough to shift a line break and cascade every element below it. Using system fonts in your theme reduces the risk, but a pixel-perfect match is not guaranteed.

---

## Requirements

- **Google Chrome** or **Microsoft Edge** — that's it. No LibreOffice, no extra runtime.

---

## Quick Start

1. Open a Marp Markdown file (`.md`) in VS Code
2. Press `F1` and run **Marp: Export to Editable PPTX**
3. Choose a save location in the dialog
4. Open the generated `.pptx` in PowerPoint — every element is editable

---

## Slide Quality

Each image shows **Marp HTML on the left** and **the exported PPTX on the right**.  
All 67 slides are from the fixture deck [`src/native-pptx/test-fixtures/pptx-export.md`](src/native-pptx/test-fixtures/pptx-export.md) and are automatically updated by CI on every release.

<!-- Screenshot comparison table — auto-updated by the Update Screenshots workflow -->

<details open>
<summary>All slide comparisons (68 slides)</summary>

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
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-061.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-062.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-063.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-064.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-065.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-066.png"></td>
</tr>
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-067.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-068.png"></td>
</tr>
</table>

</details>

---

## How it works

1. Converts your Markdown to HTML using [@marp-team/marp-cli](https://github.com/marp-team/marp-cli)
2. Launches a headless Chrome/Edge to render each slide at full resolution
3. Extracts every element's exact position, font, color, and content via `getComputedStyle()` and `getBoundingClientRect()`
4. Assembles an editable `.pptx` where each element is a native PowerPoint shape

Because it reads the browser's computed layout — not the Markdown source or CSS — it works correctly with any Marp theme, custom CSS, and `html: true` content. Element positions, sizes, colors, and images are reproduced faithfully; see the [Layout faithfulness note](#why-this-extension) above for the known typographic limitation.

---

## For contributors

See [CONTRIBUTING.md](CONTRIBUTING.md) for setup, the fix workflow (ADR log, fixture slides, visual comparison), commit style, and PR rules.

### AI-assisted development

This repository ships with [GitHub Copilot](https://github.com/features/copilot) customizations:

- **Skill** — `.github/skills/marp-pptx-visual-diff/SKILL.md`: step-by-step guide for the visual fidelity improvement loop
- **Instructions** — `.github/instructions/marp-editable-pptx.instructions.md`: coding conventions, architecture rules, and degression-prevention checklist

### Architecture & decisions

[`src/native-pptx/README.md`](src/native-pptx/README.md) contains the full architecture description and an ADR (Architecture Decision Record) log that documents every significant design decision and past bug fix.

---

## License

[MIT](LICENSE)

