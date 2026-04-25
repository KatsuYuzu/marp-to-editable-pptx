# native-pptx — Editable PPTX without LibreOffice

This module generates fully editable PowerPoint files (`.pptx`) from Marp HTML
output using browser-DOM extraction + PptxGenJS, with no dependency on
LibreOffice or any external office converter.

> **Language policy**: All source code, comments, test case names, and
> documentation in this module are written in **English**.

---

## Core design principle

**Reproduce the browser's HTML rendering as a PPTX file.**

Every layout decision, colour, font size, spacing, and image position is
resolved by the browser (Puppeteer/Chromium) using its CSS engine — not by
this code.  `getComputedStyle()` and `getBoundingClientRect()` are the sole
sources of truth for visual properties.  `slide-builder.ts` maps those
already-resolved values 1:1 to PptxGenJS primitives.

This means the module never needs to understand Marp themes, CSS selectors,
or Markdown syntax.  If the browser renders it, the module can reproduce it.
Theme changes, custom CSS, `flex`/`grid` layouts, and `html: true` content
all work automatically without any code changes here.

**Consequence for individual workarounds:**
Handling specific elements or behaviours as special cases (e.g. hard-coding
Mermaid support) is only acceptable when the browser already renders the final
output but some PPTX-side limitation prevents faithful reproduction (e.g.
`<foreignObject>` in SVG).  The fix must then be: capture the
already-rendered browser output as a raster image.  Do NOT inject rendering
logic into this module.

---

## Why this module exists (ADR)

### Background: the limits of existing PPTX export

Before this module, marp-vscode had two PPTX export code paths:

| Mode                       | Mechanism                                                                     | Limitation                                    |
| -------------------------- | ----------------------------------------------------------------------------- | --------------------------------------------- |
| **Non-editable**           | Marp HTML → Puppeteer → PNG screenshot per slide → pptxgenjs background image | Text is a bitmap — not selectable or editable |
| **Editable (LibreOffice)** | Marp HTML → Puppeteer → PDF → `soffice --headless` PDF import → PPTX          | Requires LibreOffice (experimental)           |

Some context on the history:

- **Native editable PPTX** had been [marked `wontfix`](https://github.com/marp-team/marp-cli/issues/166)
  due to visual fidelity concerns — an image-based approach was considered the
  only practical way to reproduce slide appearance accurately.
- **LibreOffice export** was later added experimentally, since its PDF import
  filter could produce a reasonably editable PPTX. It remains experimental.
- However, **enterprise users who cannot install LibreOffice** still need a way
  to get editable PPTX output. This module addresses that gap.

### Chosen approach: browser-DOM extraction

Rather than parsing the Markdown AST (no theme colour information) or
maintaining a per-theme colour DB (brittle), this module leverages the
Puppeteer browser that is **already running** for the PPTX export pipeline.

After `page.goto()`, `page.evaluate(extractSlides)` runs inside the browser
and collects:

- **`getBoundingClientRect()`** — layout-computed absolute coordinates (px)
- **`getComputedStyle()`** — resolved colour, font, weight, and text alignment
- **`textContent` / text runs** — actual rendered text with inline style spans
- **`<img>.src`** — fully-resolved image URLs including data URIs
- **`<table>` / `<ul>` / `<ol>`** — structured data extraction

The resulting `SlideData[]` is mapped to PptxGenJS API calls by `buildPptx()`
with coordinates converted from px to inches at 96 dpi.

**Key advantage**: theme-agnostic. No matter what CSS the theme uses, the
browser has already computed the final values. Custom HTML (`flex`, `grid`,
`absolute`), scoped styles, and custom themes all work transparently.

### Approaches not taken

All of the following alternatives were considered and rejected:

| Approach                         | Rejected because                                                              |
| -------------------------------- | ----------------------------------------------------------------------------- |
| A: HTML parse + theme colour DB  | Cannot compute `flex`/`grid` layout; DB requires maintenance per theme update |
| B: Markdown AST → pptxgenjs      | No theme colour information in AST; custom HTML blocks are opaque             |
| C: PNG background + text overlay | Same Puppeteer dependency but lower editability                               |
| D: Direct Open XML construction  | Re-implements what PptxGenJS already abstracts                                |
| E: PDF → pptxgenjs               | PDF text extraction quality is poor; no improvement over the LibreOffice path |

---

## Architecture

```
Markdown
  └─ marp-cli (bespoke HTML) ──────────────────────────────────┐
                                                                 │
src/native-pptx/index.ts  ◄── entry point                       │
  1. puppeteer.launch()                                          │
  2. page.goto(bespoke HTML)  ◄──────────────────────────────── ┘
  3. page.addStyleTag()   hide OSC overlay + note panels
  4. page.addScriptTag(DOM_WALKER_SCRIPT)
  5. page.evaluate(extractSlides)  ──►  SlideData[]
  6. findMissingLocalUrls()        check which local image files are accessible
     → buildBrokenContentImageJobs()  screenshot browser broken-image rendering
     → pruneMissingBackgrounds()      remove inaccessible background entries
  7. rasterizeSlideTargets()       screenshot CSS-filtered / partial images
  8. resolveImageUrls()            convert remaining local paths to data: URLs
  9. buildPptx(slides)  ──►  PptxGenJS buffer
  └──► .pptx file
```

### File map

| File                             | Role                                                                                                                         |
| -------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `index.ts`                       | Pipeline orchestration: browser launch → DOM extract → rasterize → build PPTX                                                |
| `dom-walker.ts`                  | `extractSlides()` — runs **in the browser via `page.evaluate()`**; reads DOM, returns `SlideData[]`                          |
| `dom-walker-script.generated.ts` | Compiled IIFE string of `dom-walker.ts` — injected via `addScriptTag` (regenerate with `node src/native-pptx/scripts/generate-dom-walker-script.js`) |
| `slide-builder.ts`               | `buildPptx()` + `placeElement()` — maps `SlideData[]` to PptxGenJS API calls                                                 |
| `types.ts`                       | Shared TypeScript types (`SlideData`, `SlideElement`, `TextRun`, …)                                                          |
| `browser.ts`                     | Chrome/Chromium auto-detection utilities                                                                                     |
| `utils.ts`                       | `pxToInches`, `rgbToHex`, `pxToPoints`, `cleanFontFamily`, `sanitizeText`                                                    |

### Why `dom-walker-script.generated.ts`?

`page.evaluate(fn)` serialises `fn.toString()`. When webpack/esbuild minimises
the bundle it injects module-scope helpers (e.g. `t(fn, name)`) into function
bodies; after serialisation those references are undefined in the browser
context and cause `ReferenceError`.

To avoid this, `dom-walker.ts` is compiled separately by esbuild into a
**standalone IIFE** (`src/native-pptx/scripts/generate-dom-walker-script.js`) and embedded as
a string constant. `page.addScriptTag({ content: DOM_WALKER_SCRIPT })` injects
it safely.

> **Important**: after any change to `dom-walker.ts`, run
> `node src/native-pptx/scripts/generate-dom-walker-script.js` to update the generated file.

---

## Key implementation details

### Presenter notes

marp-cli bespoke HTML injects notes as:

```html
<div class="bespoke-marp-note" data-index="0" tabindex="0">
  <p>Note text</p>
</div>
```

`data-index` is the zero-based slide index. `dom-walker.ts` falls back to this
selector when the raw-Marpit `[data-marpit-presenter-notes]` attribute is absent
(which is always the case for marp-cli bespoke HTML output).

### Image rasterization and missing-image handling

Several categories of images cannot be faithfully embedded as-is into a PPTX
file and require special treatment.  All such images are handled by taking a
Puppeteer screenshot and embedding the resulting PNG data URL instead.

The following passes run in `index.ts` **before `buildPptx()`**:

1. `buildBrokenContentImageJobs` — screenshot the browser's own broken-image
   indicator (icon + alt/filename) for content images whose source file is
   missing on disk.  Produces the same visual shown in HTML and PDF output.
2. `buildFilteredBgJobs` — `<figure>` background images with CSS filters
   (`grayscale`, `blur`, `brightness`, …).
3. `buildCssFallbackBgJobs` — CSS `background-image` set by Marp directives
   (captured as a full-slide screenshot).
4. `buildFilteredContentImageJobs` — inline `<img>` elements with CSS filters.
5. `buildRasterizeImageJobs` — images explicitly flagged `rasterize: true`
   (e.g. Mermaid SVGs with `<foreignObject>`).
6. `buildPartialBgJobs` — partial-width background images (`![bg right:30%]`)
   where CSS `background-size: cover` crops differently than PPTX stretch-to-fill.

Missing background images (`![bg](missing.png)`) are handled separately by
`pruneMissingBackgrounds()`, which removes them from `slide.backgroundImages`
so the slide falls back to its solid background-color fill — matching CSS
behaviour (no broken-image indicator for backgrounds).

Before screenshots are taken, the bespoke **OSC overlay**
(`<div class="bespoke-marp-osc">`) and note panels (`.bespoke-marp-note`) are
hidden via `page.addStyleTag()` so they do not appear in the output.

### Mermaid diagrams

When `--html` is passed to marp-cli, `<div class="mermaid">` + `<script src="mermaid.min.js">` in the Markdown are output to HTML as-is. Puppeteer waits for the CDN script load with `waitUntil: 'networkidle0'`, then after a 1-second settle delay, mermaid's async rendering (Promise-based) completes. The DOM walker then converts the resulting `<svg>` to a data URL and captures it.

mermaid@10 renders flowchart text labels via `<foreignObject>`. PowerPoint cannot natively render `<foreignObject>` as part of an SVG, so the DOM walker sets `rasterize: true` when a `<foreignObject>` exists inside an `<svg>`, and `index.ts` replaces those elements with PNG screenshots.

**Prerequisite**: when the extension setting `markdown.marp.html` is `"all"`, `--html` is automatically forwarded to marp-cli (controlled in `extension.ts`). When it is not `"all"`, marp-cli's security policy strips script tags and mermaid does not work — this is intentional (same behaviour as marp for VS Code).
### Text height clamping

Font rendering differences between Chromium and PowerPoint can cause text
elements near the bottom of a slide to have a computed height that extends
beyond the slide boundary. `placeElement()` in `slide-builder.ts` clamps the
height of all text-type elements so `y + h ≤ slideH`. Images are **not**
clamped — overflow is intentional for split-layout backgrounds.

### Heading border-left text offset

When a heading has `border-left` (common in themes as a vertical accent bar),
`slide-builder.ts` draws the colour rectangle **before** the text box and shifts
the text box right by the border width (`x + bw`, `w - bw`). This mirrors the
same pattern used for `<blockquote>`, ensuring the decorative bar is _behind_
the text in the PPTX z-order and that text does not visually overlap the bar.

### ZWJ emoji sequence preservation

`sanitizeText` strips control and zero-width characters to prevent PptxGenJS
from emitting corrupt XML. **U+200D (Zero Width Joiner) is intentionally
preserved** because it is load-bearing for multi-codepoint emoji compositions
(e.g. `U+1F9D1 U+200D U+1F4BB` → 🧑‍💻). Stripping it would split the sequence
into two separate glyphs (🧑 + 💻).

### Leading-badge heading offset (`computeLeadingOffset`)

Inline badge shapes (`.step`, `.badge-current`, etc.) are extracted as separate
PptxGenJS shapes. When a badge sits at the _left edge_ of a heading or paragraph
box ("leading badge"), the text box is shifted right by the badge's width so
that text does not render on top of the shape. `computeLeadingOffset` computes
this offset by finding badge shapes whose `x` is within 8px of the container's
left edge.

### Slide numbers

Marp renders page numbers as an HTML/CSS pseudo-element (`section::after`).
That approach is visually fragile in PPTX export because the pseudo-element can
be restyled or suppressed by unrelated theme rules (`::after` banners,
decorative bars, split-background layers, etc.). It also produces fixed text,
which does not renumber automatically after slide reordering inside PowerPoint.

This module therefore treats pagination differently from ordinary content:

- `dom-walker.ts` keeps only the raw `data-marpit-pagination` presence as a
  deck-wide source flag (`sourceHasPagination`)
- `slide-builder.ts` does **not** emit a page-number text element from HTML
- the PPTX builder enables PowerPoint's built-in slide-number field
  consistently for the whole deck when pagination is used at all
- `dom-walker.ts` keeps decorative pagination pseudo-element backgrounds as
  shapes when they provide visible bars / ribbons / pills, while leaving the
  page number itself to PowerPoint's native slide-number field
- `index.ts` makes only the HTML pagination text transparent during
  rasterization so screenshot-based backgrounds do not burn in duplicate page
  numbers while decorative pseudo-element backgrounds remain visible

This is an intentional exception to the usual "browser rendering is the source
of truth" rule. For slide numbers, editability and correct renumbering in
PowerPoint are more important than reproducing the exact HTML pseudo-element.
Unlike ordinary content, slide numbers are treated as deck metadata rather than
per-slide layout that should be reconstructed from CSS.

---

## Supported elements

| Element                               | Fidelity | Notes                                         |
| ------------------------------------- | -------- | --------------------------------------------- |
| Slide background (solid colour)       | ◎        | Extracted via `getComputedStyle`              |
| Slide background (image / CSS filter) | ◎        | Rasterized by Puppeteer                       |
| Heading H1–H6                         | ◎        | Inline run styling, border-bottom/left        |
| Paragraph                             | ◎        | Multiple runs with bold/italic/underline/link |
| Bulleted / numbered list              | ◎        | Nested lists, tight-list emoji bullets        |
| Table                                 | ◎        | Per-cell style, colour, alignment             |
| Code block                            | ○        | Syntax-highlighted runs preserved             |
| Image (URL / data URI / file://)      | ◎        | Natural size with aspect ratio                |
| Mermaid diagram (SVG)                 | ◎        | Rasterized to PNG                             |
| Blockquote                            | ○        | Left border bar + text                        |
| Header / Footer                       | ◎        | Absolute coordinate placement                 |
| Presenter notes                       | ◎        | Both raw-Marpit and bespoke-HTML formats      |
| CSS gradient background               | △        | Simplified to solid colour                    |
| CSS `transform` / `clip-path`         | △        | Ignored; elements placed at rect coordinates  |

---

## Running the visual diff improvement loop

This workflow compares rendered HTML slides against PPTX output to identify
remaining fidelity gaps.

### Canonical test deck

`src/native-pptx/test-fixtures/pptx-export.md` is the primary edge-case reference for
this module. It contains 63 slides covering every known rendering challenge:

- Basic headings, paragraphs, lists, tables, code blocks
- `border-bottom` on H1 and `border-left` vertical bar on H2/H3
- Inline badge shapes (pill, circle, status chips, step numbers)
- Leading-badge heading offset (`computeLeadingOffset`)
- ZWJ emoji sequences (🧑‍💻, 👨‍👩‍👧‍👦 — multi-codepoint single glyph)
- `strong { background-color: ... }` solid colour highlight
- CSS `section::before/after` banner suppression (scoped suppression + global-rule suppression)
- Complex background filters (`blur`, `brightness`, `grayscale`)
- Split-layout (`flex`, `grid`, HTML `<div>`) with images

Use this file as the input when running the visual diff loop below. The
file lives inside the repository so any contributor can reproduce results.

### Prerequisites

1. Install dependencies: `npm install`
2. Build the native-pptx bundle:
  `node src/native-pptx/scripts/generate-dom-walker-script.js`
  `node src/native-pptx/scripts/build-native-pptx-bundle.js`
3. Have Chrome/Chromium available (auto-detected via `@puppeteer/browsers`)
4. Install LibreOffice for PPTX → PNG rendering, **or** use PowerPoint COM
   automation on Windows

### Steps

```sh
# 1. Generate the canonical bespoke HTML (--html is required for card/mermaid slides)
npx marp src/native-pptx/test-fixtures/pptx-export.md \
  --html --allow-local-files \
  --output src/native-pptx/test-fixtures/slides-ci.html

# 2. Generate native PPTX from the HTML
node src/native-pptx/tools/gen-pptx.js \
  src/native-pptx/test-fixtures/slides-ci.html \
  dist/compare-out.pptx

# 3. Compare HTML slides vs PPTX slides side-by-side
#    (requires PowerPoint COM on Windows, or LibreOffice on Linux/macOS)
node src/native-pptx/tools/compare-visuals.js \
  src/native-pptx/test-fixtures/slides-ci.html \
  dist/compare-out.pptx

# Output always goes to dist/compare-slides-ci/ (never inside src/):
#   html-slide-NNN.png   — Marp HTML reference screenshot
#   pptx-slide-NNN.png   — PPTX slide screenshot
#   compare-NNN.png      — side-by-side diff image
#   compare-report.html  — per-slide diff area summary
```

### Debug mode

Set `MARP_PPTX_DEBUG=1` before running `gen-pptx.js` to dump a
`*.native-pptx.json` file alongside the output, containing the full
`SlideData[]` extracted from the DOM:

```sh
MARP_PPTX_DEBUG=1 node src/native-pptx/tools/gen-pptx.js docs/example.html
```

### AI-assisted improvement loop

1. Run `compare-visuals.js` and review `compare-report.html`.
2. Open a high-diff slide pair (`html-slide-NNN.png` + `pptx-slide-NNN.png`).
3. Describe the visual delta to an AI agent with access to this codebase.
4. The agent locates the responsible code in `dom-walker.ts` or
   `slide-builder.ts`, proposes a fix, and updates the unit tests.
5. If `dom-walker.ts` changed, regenerate `dom-walker-script.generated.ts`.
6. Re-run `gen-pptx.js` + `compare-visuals.js` to verify the improvement.

Repeat until all slides in `compare-report.html` show acceptable diff scores.

---

## Running tests

```sh
# Unit tests only (fast — no browser required)
npx jest "native-pptx"

# Full test suite
npx jest
```

### Test file overview

| Test file               | What it covers                                                                                                                                 |
| ----------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| `index.test.ts`         | Pipeline orchestration: browser lifecycle, CSS injection, script injection, evaluate call, buffer return                                       |
| `dom-walker.test.ts`    | `extractSlides()`: slides from SVG/inline HTML, backgrounds, element classification, tables, lists, presenter notes (both formats)             |
| `slide-builder.test.ts` | `buildPptx()` / `placeElement()`: heading/paragraph/list/table/code/image/blockquote placement, text-height clamping, image overflow exemption |
| `browser.test.ts`       | `findChrome()`: platform detection, path resolution                                                                                            |
| `utils.test.ts`         | `pxToInches`, `rgbToHex`, `pxToPoints`, `cleanFontFamily`, `sanitizeText`                                                                      |

---

## Known limitations

- **CSS gradients**: pptxgenjs gradient API differs from CSS; gradients are
  simplified to a solid colour.
- **CSS `transform`**: rotated/scaled elements are placed at their unrotated
  bounding-box coordinates.
- **`clip-path`**: ignored; elements appear at full size.
- **Text overflow at runtime**: text that exceeds a box inside PowerPoint is
  hidden by PPTX's own clipping. Only the height measured at extraction time
  is clamped.
- **Web fonts**: PowerPoint falls back to an installed font when the web font
  is not embedded. `cleanFontFamily()` strips CSS font-stack fallbacks and
  keeps only the primary family name.
- **Dark mode / forced-colors**: screenshots are taken in light mode.

---

## Bug fix and decision log (ADR log)

This section is a decision log for ensuring reproducibility in AI-assisted development.
Re-writing code without understanding why things are the way they are risks reproducing the same bugs,
so every discovered problem, its root cause, and its resolution are recorded here.

All ADR entries are written in English.

### ADR-01: toListTextProps did not output highlight (slide 56/58)

**Problem**
When `<strong style="background-color:#f1c40f">` was inside a `<li>`, the PPTX list-item text did not have the highlight colour applied.

**Root cause**
`toListTextProps` in `slide-builder.ts` was missing the code to convert `run.backgroundColor` to `highlight`. `toTextProps` (for paragraphs) had it implemented correctly, but `toListTextProps` was overlooked during development.

**DOM walker side was correct**
When `extractListItems` calls `extractTextRuns(<strong>)`, `extractTextRuns` propagates the element's own `elementBg` (`rgb(241, 196, 15)`) as `bg` for TEXT_NODEs, so `run.backgroundColor` was set correctly. The problem existed only on the builder side.

**Fix**
Added `highlight: run.backgroundColor ? rgbToHex(run.backgroundColor) : undefined` to the run mapping in `toListTextProps`.

**Why it was not caught**
The tests for `toListTextProps` did not include any run with `backgroundColor`. Tests for `highlight` in `toTextProps` existed, but no equivalent test was written for `toListTextProps`.

**Regression prevention**
Added test: `toListTextProps — run with backgroundColor gets highlight set`. This detects the regression at build time.

---

### ADR-02: mermaid SVG foreignObject was not rendered in PowerPoint

**Problem**
mermaid@10 flowcharts appeared in PPTX but text labels and arrows were missing.

**Root cause**
mermaid@10 renders flowchart node text using `<foreignObject><div>...</div></foreignObject>`. PowerPoint ignores the `<foreignObject>` content when embedding SVGs, causing text to disappear.

The DOM walker's `tag === 'svg'` branch base64-encoded the entire SVG into a data URL but did not set `rasterize: true`. The `tag === 'pre'` branch (code-fence mermaid) already set `rasterize: true`, creating an inconsistency.

**Fix**
In the `tag === 'svg'` branch, check `child.querySelector('foreignObject') !== null`. If true, set `rasterize: true`. `buildRasterizeImageJobs` in `index.ts` replaces these elements with Puppeteer PNG screenshots.

**Why it was not caught**
The problem was only visible in CI comparison images (`compare-NNN.png`). No unit test covered SVG+foreignObject, and the CI visual diff had no threshold-based alerting.

**Regression prevention**
1. Added per-slide RMSE report to CI (outputs WARN when RMSE > 0.20).
2. A test for `<svg>` + `<foreignObject>` yielding `rasterize: true` in `dom-walker.test.ts` is desirable (TODO).

---

### ADR-03: mermaid must use div.mermaid syntax, not code fence

**Problem**
The test fixture used ` ```mermaid ` code fence syntax but PPTX conversion failed.

**Root cause**
marp-core 4.x does not have a mermaid transform plugin that processes ` ```mermaid `. The code fence is output as `<pre><code class="language-mermaid">` and is never converted to SVG.

**Correct approach**
`<div class="mermaid">diagram code</div>` is the correct syntax. mermaid.js (loaded from CDN) transforms this div into SVG. The script tag only needs to appear once in the same slide or a preceding slide.

**Prerequisite**
`html: true` in frontmatter, and VS Code setting `markdown.marp.html: "all"`. With just `html: true`, marp-cli strips `<script>` tags for security policy reasons.

---

### ADR-04: Output the temporary HTML to the same directory as the source MD

**Problem**
Relative-path images (`./attachments/image.png` etc.) did not appear in the PPTX.

**Root cause**
Outputting the temporary HTML to `os.tmpdir()` caused marp-cli to resolve image relative paths relative to the temp directory, so images next to the original MD file were not found.

**Fix**
Changed to output the temporary HTML to `path.dirname(doc.uri.fsPath)`. The extension deletes it during cleanup.

---

### ADR-05: CI visual regression test design

**Current state**
CI (`screenshots.yml`) generates side-by-side compare images of HTML screenshots vs PPTX screenshots and publishes them to gh-pages. Human visual inspection.

**Problem**
Visual inspection produces missed regressions (the highlight bug in ADR-01 was not noticed for an extended period).

**Decided improvements**
1. **Per-slide RMSE report**: Added ImageMagick `compare -metric RMSE` step to CI. Slides with RMSE > 0.20 are output as WARN in the CI log. This step uses `continue-on-error: true` to avoid failing the job (HTML vs PPTX always has some difference).
2. **Unit tests to guarantee structure**: The root cause of visual problems is often a `SlideData[]` structure issue (e.g. missing `backgroundColor`). Unit tests that strictly verify `SlideData[]` structure enable browser-free detection before CI.
3. **63-slide test fixture**: `pptx-export.md` is a canonical fixture covering all known edge cases. When a new bug is discovered, a corresponding slide must be added to this file before fixing.

**Future work (TODO)**
- Add PPTX-to-PPTX regression (compare across commits) to detect pure PPTX rendering regressions.
- Add `pixelmatch` diff images highlighting differences in red (`pngjs` + `pixelmatch` already in devDependencies).

---

### ADR-06: Repository script and output directory structure

**Directories and their purposes:**

| Path | Purpose | Kept in git? |
|---|---|---|
| `scripts/` | Project-level tooling: extension build (`copy-marp-cli-assets.js`) and CI screenshot generation (`gen-html-screenshots.js`) | Yes (source) |
| `src/native-pptx/scripts/` | Module build tools: generate the DOM walker IIFE string (`generate-dom-walker-script.js`) and bundle the standalone CJS module (`build-native-pptx-bundle.js`) | Yes (source) |
| `src/native-pptx/tools/` | Developer debugging tools: `gen-pptx.js`, `compare-visuals.js`, `md-to-html.js`, `diagnose-pptx.js` | Yes (source) |
| `lib/native-pptx.cjs` | Build output of `build-native-pptx-bundle.js`. Used by the developer tools above. | **No** (gitignored) |
| `dist/` | All build/comparison artifacts. Never inside src/. | **No** (gitignored) |

**Why `compare-visuals.js` always outputs to `dist/`:**
The tool accepts an arbitrary HTML input path, so placing output next to the input would scatter artifacts through `src/`. Using `path.resolve(__dirname, '../../..')` to find the project root and writing to `dist/compare-<basename>/` keeps all generated files in the gitignored `dist/` directory regardless of where the input HTML lives.

**Why `scripts/` and `src/native-pptx/scripts/` are not merged:**
Top-level `scripts/` contains tooling for the extension project as a whole (webpack pipeline, CI).
`src/native-pptx/scripts/` contains build tooling specific to the native-pptx module (code generation, bundling).
Merging these would create confusion about which tools affect the VS Code extension build vs the standalone module build.

**Why `lib/native-pptx.cjs` is gitignored:**
It is a compiled artifact produced from TypeScript source. Committing it would mean the file goes stale on every source change. Developers and CI regenerate it on demand with `node src/native-pptx/scripts/build-native-pptx-bundle.js`.

---

### ADR-07: Duplicate `breakLine` insertion from `<br>` + newline text node

**Problem**
Extra blank lines appeared between bullets and around emojis in slides p4 (emoji), p18 (step badge), and p19 (number badge).

**Root cause**
The whitespace-only TEXT_NODE handling in `extractTextRuns` (breakLine insertion by counting `\n`) was double-counting the `\n` text node that immediately follows a `<br>` tag.

```
<p>✅ Sentence<br>\n🚀 Next</p>
           ^^^ tag → breakLine
              ^^^ text node \n → breakLine again
```

**Fix**
Added a `lastIsBreak()` guard that suppresses duplicates when the previous item is already a `breakLine`. Located in the TEXT_NODE branch loop inside `extractTextRuns` in `dom-walker.ts`.

---

### ADR-08: Paragraph placement when `<img>` is the first child of a `<p>`

**Problem**
With `![w:300](img.png) Normal image beside text`, text was placed overlapping the image at the top-right. With `![w:300](img.png) \n Caption`, text was placed at the top of the image.

**Root cause**
The `<p>` branch in `walkElements` used the paragraph's full `getBoundingClientRect()` rect as the text box. When `<img>` was the first child, the paragraph rect started at `(image.left, image.top)`, so text and image completely overlapped.

**Fix (2 patterns)**

| Pattern | HTML structure | PPTX conversion |
|---|---|---|
| Case A (side by side) | `<img> text` (no newline) | `paragraph.x += imgWidth`, `paragraph.y += imgHeight - lineHeight` (bottom-right of image) |
| Case B (stacked) | `<img><br>text` | `paragraph.y += imgHeight`, `paragraph.height -= imgHeight` (directly below image) |

Case A y-offset: CSS `vertical-align: baseline` aligns the text baseline to the image bottom, so `inlineImgYOffset = max(0, imgBottom - lineHeight)` moves the paragraph y down to near the image bottom.

**Detection logic**
Walk the `<p>` childNodes from the front, recording the first non-emoji `<img>`. If a `<br>` immediately follows it, Case B; otherwise Case A.

---

### ADR-09: Table row background colour could not be retrieved

**Problem**
Alternating row colours did not appear in PPTX for page 7 (Table) and page 25 (Complex cells).

**Root cause**
Marp's CSS applies `:nth-child(even)` rules to `<tr>`. Browsers do not inherit `<tr>` colour into `<td>` via `getComputedStyle().backgroundColor` (it remains `transparent`).

```
tr:nth-child(even) { background: rgb(246, 248, 250) }
td → getComputedStyle(td).backgroundColor = rgba(0,0,0,0)  ← not inherited
```

**Fix**
`extractTableData` now also reads `<tr>` backgroundColor and applies it as a fallback when `<td>` is transparent.

---

### ADR-10: Headings wrapped in PPTX due to font metrics differences

**Problem**
Single-line headings in HTML wrapped to two lines in PPTX on pages 27, 41, 45, and others.

**Root cause**
Chrome (Skia) and PowerPoint (DirectWrite) measure fonts slightly differently. DirectWrite measures characters slightly wider at the same point size, causing lines that fit in HTML to overflow and wrap in PPTX. The text box width uses `offsetWidth` from HTML, leaving no pixel margin.

**Fix**
Full-width headings (where `el.x < slideW * 0.15` and `el.x + el.width > slideW * 0.85`) have their text box width extended to `slideW - el.x - 16` px. The 16 px serves as a right-edge margin buffer.

**Why there are no side effects**
Widening the heading box does not cause overlap with adjacent elements (Marp headings are basically text-only with no adjacent elements). If the text is short it simply fits within the extended width with no visual change.

---

### ADR-11: compare-visuals.js incorrectly counted `![bg]` slides as MISSING

**Problem**
Comparing `slides-ci.html` (59 slides) vs PPTX (59 slides) reported HTML: 56, marking slides 57-59 as MISSING.
(Currently 63 slides; slides 62 and 63 were added after the ADR-16 fix.)

**Root cause**
When Marp outputs slides with `![bg]` in "advanced background" mode, each slide is split into 3 `<section>` layers: `background` (background image), `content` (slide content), and `pseudo` (page number). The previous count logic only counted `<section>` elements without the `data-marpit-advanced-background` attribute, missing the `content` layer sections and producing a lower total.

**Fix**
Grouped all sections by `data-marpit-pagination` key, and counted unique keys excluding the `pseudo` layer as the slide count. This is the same approach as the PPTX exporter's `slideGroups` logic, ensuring the two always agree.

**Result**
HTML: 59, PPTX: 59, MISSING: 0 can now be consistently confirmed.

---

### ADR-12: `--html` flag required to regenerate `slides-ci.html`

**Problem**
After regenerating `slides-ci.html` without the `--html` flag, cards (div-based), badges, and mermaid diagrams were missing.

**Root cause**
marp-cli strips HTML blocks (`<div>`, `<span>`, `<script>`, etc.) by default for security policy reasons. They are only preserved in the HTML output when the `--html` flag is passed.

**Correct regeneration command**
```sh
npx marp pptx-export.md --html --allow-local-files --output slides-ci.html
```

`--allow-local-files` is required to allow image loading via the `file://` protocol. Without it, relative-path images return 403.

**Note**
`slides-ci.html` is excluded by `.gitignore`. CI uses `scripts/gen-html-screenshots.js` which passes the flags correctly. Always include `--html` when manually regenerating locally.

---

### ADR-13: Global `section::before/::after` generated spurious bars on user-classed slides

**Problem**
Converting a Marp slide with global theme CSS such as `section::before { content: ""; background: #16324f; height: 16px; }` (applied to all slides) produced dark blue bars at the top of **only slides with user classes** like `_class: cover` or `_class: agenda`. Slides without a class showed no bar, creating a visual inconsistency.

**Root cause**
`extractPseudoElements` had the following logic for pseudo-elements with `content: ''`:

> "Section has a user class → extract it (expected to be a class-specific decorator like `section.decorated::before`)"
> "No class → skip (expected to be a Marp scoped artifact)"

However, when the global rule `section::before` is scoped by Marp it becomes `section[data-marpit-scope-XXX]::before`, which applies the same background colour to all sections regardless of user class. With the "extract if has user class" logic, only classed slides produced the bar.

**Fix (`dom-walker.ts`)**
After building slide groups, collect the background colours of `content: ''` pseudo-elements on classless sections into `globalPseudoSignatures` (a `Set<string>`). In `extractPseudoElements`, when a pseudo-element has `content: ''` and the section has a user class, **skip extraction if the same background colour is in the global signatures**.

```typescript
// Collection phase (classless sections only)
for (const { content } of slideGroups.values()) {
  if (!content || (content as HTMLElement).className?.trim()) continue
  for (const pseudo of ['::before', '::after'] as const) {
    const ps = getComputedStyle(content, pseudo)
    // content:'' and opaque background → register as global signature
    ...
    globalPseudoSignatures.add(`${pseudo}:${bg}`)
  }
}

// In extractPseudoElements
if (stripped === '') {
  if (!sectionClass) continue                          // existing: skip when no class
  if (globalPseudoSignatures.has(`${pseudo}:${pgBg}`)) continue  // new: skip global rules
}
```

**Preserved behaviours**

| Case | Result |
|---|---|
| `section.decorated::before` only — classless section is transparent | Extracts bar on decorated slides ✓ |
| Global `section::before` — all slides same colour | Skipped on all slides ✓ |
| `section.agenda::after { background: orange }` — classless has blue | Extracts orange bar on agenda only ✓ |

**Tests added (`dom-walker.test.ts`)**
- `global section::before (same colour as classless section) — suppressed even on classed slides`
- `class-specific decorator with different colour than classless section — still extracted`

**Fixture added (`pptx-export.md` Slide 60)**
Added a slide combining `_class: decorated` and scoped `section::before/::after` so that absence of the spurious bar can be confirmed in the visual comparison report.

---

### ADR-14: Handling missing image files (transparent PNG → screenshot)

**Problem**
Converting a slide with a reference to a non-existent local image (`![](missing.png)`) threw an `ENOENT` exception synchronously inside `pptx.write()` via `fs.readFileSync` in PptxGenJS, showing a VS Code error dialog.

**Background (interim fix before ADR-14)**
The first fix replaced missing files with a 1x1 transparent PNG placeholder. This suppressed the error but caused PowerPoint to render the transparent PNG as a mysterious green shape. Regardless of the cause, the approach was judged inappropriate because it does not communicate to the user that an image is missing.

**Problem classification**

| Type | Correct behaviour | Reason |
|---|---|---|
| Content image `![](missing.png)` | Embed the browser's broken-image icon + alt/filename as PNG | Same appearance as HTML/PDF output. User must be notified of the missing file. |
| Background image `![bg](missing.png)` | Remove from `backgroundImages` and fall back to slide background colour | CSS `background-image` ignores non-existent URLs — no error indicator. |

**Decided implementation**
Added 3 steps immediately after DOM extraction in `generateNativePptx` (before other rasterize passes):

1. `findMissingLocalUrls(slides)` — run `fs.access` on all local-path image URLs across all slides, collecting inaccessible URLs.
2. `buildBrokenContentImageJobs(slides, missingUrls)` — take a Puppeteer screenshot of the element region (`img.x/y/width/height`) of missing content images and overwrite `img.src` with the data URL. Chromium has already rendered its own broken-image icon for `<img>` elements with non-existent `file:///` paths.
3. `pruneMissingBackgrounds(slides, missingUrls)` — remove background entries with missing URLs from the `backgroundImages` array.

The silent ENOENT fallback in `fileUrlToDataUrl` was removed. Missing files are handled by these 3 steps; any missing file reaching `resolveImageUrls` indicates an unexpected I/O failure that should surface via the error dialog.

**Tests added (`index.test.ts`)**
- `when a content image is missing, screenshot its region and replace src`
- `when a background image is missing, remove from backgroundImages and fall back to slide background colour`

---

### ADR-15: Direct text nodes in flex/grid containers were lost when blockChildren existed

**Problem**
A table-of-contents slide using `html: true` + a custom class (e.g. `.agenda-item` — a `display:flex` div containing an `inline-flex` badge span + a direct text node "Agenda topic") produced a PPTX where only the badge number appeared; the adjacent agenda text was completely absent.

**Root cause: mixed responsibility between `element.children` and `element.childNodes`**

`walkElements` iterates `Array.from(parent.children)` (Element nodes only). `extractTextRuns` iterates `Array.from(element.childNodes)` (includes Text nodes).

Normal divs (`display:block` or `display:flex` with all Element children) are fine. The problem arose when a container held **both Element children (→ blockChildren) AND direct TEXT_NODE children**.

```
<div class="agenda-item" style="display:flex">
  <span class="num" style="display:inline-flex; background:...">1</span>  ← Element → blockChild
  Agenda text  ← Text node → invisible to walkElements
</div>
```

Processing flow (before fix):

1. `walkElements(agenda-item)` processes `span.num` (inline-flex) in the `else` branch → `blockChildren = walkElements(span.num)` → inline-only → `extractTextRuns(span)` → `paragraph("1")` added to `blockChildren`.
2. `blockChildren.length > 0` → outputs `container { children: [paragraph("1")] }` and ends the pass.
3. TEXT_NODE `"Agenda text"` is not in `parent.children` and is **never walked**.

**Why existing tests did not catch this**
Test cases assumed either "container = all block-level children" or "container = inline-only text". No test covered "blockChildren-generating Element AND direct TEXT_NODE coexist in the same container". This pattern only occurs with `html: true` + custom CSS class + handcrafted HTML TOC components.

**Fix (`dom-walker.ts` — the `else` branch of `walkElements`)**
After outputting the container when `blockChildren.length > 0`, added a shallow walk of the container's `childNodes` that only picks up **text nodes and inline-level elements**. Block-level child elements are explicitly skipped (already handled in `blockChildren`) to prevent text duplication.

**Tests added (`dom-walker.test.ts`)**
- `text node sibling to badge span inside flex item is NOT lost`
- `two-column grid with badge+text items — all item texts extracted`

**Fixture added (`pptx-export.md` Slide 61)**
Added a TOC slide with a `display:grid` wrapper and 6 `display:flex` agenda items (each containing an `inline-flex` badge span + direct text node). Visual comparison shows text completely absent before the fix and correctly rendered after.

**Overlap fix (2nd commit)**
Fixing text node extraction alone left the paragraph x starting at the container left edge (same position as the badge), causing visual overlap. Applied `computeLeadingOffset(blockChildren, ...)` to offset the paragraph x to the right edge of the leading blockChild.

---

### ADR-K01: Latent bugs similar to ADR-15 (observation needed)

The following were found during the root cause investigation of ADR-15. They have not surfaced yet but are structurally similar risks. Their conditions are documented here as grounds for prioritisation decisions.

#### K01-A: `skipInlineBadges` is not forwarded in the block-level recursion (medium risk)

**Code**: `extractTextRuns`, block branch (line ~225)

```typescript
if (/^(block|flex|grid|list-item|table)/.test(elStyle.display)) {
  runs.push(...extractTextRuns(el))  // <- skipInlineBadges not forwarded
}
```

The inline branch correctly forwards `extractTextRuns(el, skipInlineBadges)` but the block branch always recurses with `false`.

**Impact scenario**: `html: true` + a heading or paragraph containing a `display:flex` div with a badge span inside. `extractInlineBadgeShapes` extracts the badge at the parent (heading) level, but `extractTextRuns` block recursion processes badge text with `skipInlineBadges=false`, causing the badge to appear both as a shape and as a `backgroundColor` run — duplicated.

**Condition**: `html: true` + `display:flex` container directly inside `<h1>` / `<p>` with a badge span. Does not occur in standard Marp Markdown.

**Priority**: Low (edge case requiring `html: true` + flex container inside heading).

#### K01-B: `extractListItems` does not extract badges as shapes (low risk — by design)

`extractInlineBadgeShapes` is only called in the `h1-6`, `p`, `blockquote`, `header`, and `footer` branches of `walkElements`. It is not called inside `extractListItems`.

**Impact**: With `html: true`, using `display:flex` + badge span inside `<li>` results in the badge being output as a `backgroundColor` highlight on a text run rather than as an independent shape. Visually close but `border-radius` of the circular badge is not reproduced.

**Priority**: Accepted as current behaviour. If badge shape accuracy matters, address separately.

---

### ADR-16: `<br>` missing inside `<li>` + image-between-list-items misalignment

#### Problem

1. **Trailing-space line break missing inside `<li>`**
   Markdown "two trailing spaces + newline" (`  \n`) becomes `<br>` in HTML. In `<p>` and `<h1-6>` this `<br>` was correctly converted to a `breakLine` run in PPTX. Inside `<li>` it was ignored, removing the line break.

2. **Image-between-list-items misalignment**
   Placing a list item, an image, and a list item in Markdown without blank lines causes markdown-it to parse the image as a lazy continuation of the first `<li>`. As a result the `<ul>` bounding box expanded to include the image area, and PPTX generated a single text box for it, causing the `extractNestedImages`-appended image to land at the same y coordinate as "Second item".

#### Root cause

**Bug 1**
The `else` branch of the `<li>` child-node processing loop in `extractListItems` passed `<br>` to `extractTextRuns(<br>)` as the **root element**. `<br>` has no child nodes, so `extractTextRuns` returned `[]` and no `breakLine` run was generated.

**Bug 2**
The `<ul>/<ol>` branch in `walkElements` output the entire `<ul>` as a single text box, then extracted and appended images via `extractNestedImages`. Because the `<ul>` height included the image area, the appended image landed at a y coordinate overlapping the text box.

#### Fix

**Fix 1 (`extractListItemEl` helper extraction)**
Extracted the `<li>` processing loop from `extractListItems` into `extractListItemEl(li, level)`. Added an explicit branch for `<br>`:

```typescript
} else if (childTag === 'br') {
  if (!lastIsBreak()) runs.push({ text: '', breakLine: true })
}
```

**Fix 2 (split path in `walkElements`)**
When a `<ul>/<ol>` contains a `<li>` with a non-emoji `<img>`, split the list at `<li>` boundaries containing images using a "split path". Each sub-list's y/height uses the actual `<li>.getBoundingClientRect()`. Images are inserted inline between sub-lists at actual rendered coordinates.

#### Tests added

- `dom-walker.test.ts`: 3 added
  - `extractListItems — <br> inside tight list <li>`
  - `image embedded between list items (via extractSlides)` x 2
- `slide-builder.test.ts`: 2 added (191 total)
  - `<br> continuation line uses invisible bullet to align indent`
  - `multiple continuation lines all use invisible bullet to align indent`
- Fixture: Added Slide 62 (trailing-space line break) and Slide 63 (image between list items) to `pptx-export.md`.

#### ADR-16 supplement: `toListTextProps` indent fix

Fix 1 in `dom-walker.ts` correctly generated `breakLine: true` runs. However, `slide-builder.ts`'s `toListTextProps` wrongly assumed `breakLine: true` was output as `<a:br/>`. Investigation revealed the actual PptxGenJS behaviour:

**Actual PptxGenJS behaviour (found during ADR-16 investigation)**

| Method | OOXML output |
|---|---|
| `breakLine: true` on a run | Closes the current paragraph and **starts a new `<a:p>`** |
| Next run with truthy `bullet` + no `opts.align` | `else if (bullet)` branch starts a new `<a:p>` |
| Next run with truthy `bullet` + `opts.align` set | **Never reaches `else if (bullet)`** — the `if` for align fires but without an align change the paragraph is not split |

`list`'s `addText` call always sets `align: el.style.textAlign`, so the `bullet` paragraph-trigger does not work. Continuation lines must also use `breakLine: true` to close paragraphs.

**Fix 3 (`toListTextProps` continuation-line indent fix)**
Split `item.runs` into groups at each `breakLine` run. The last run of each group gets `breakLine: true` to close the paragraph. Continuation groups (g > 0) get `bullet: { characterCode: '200B' }` (U+200B zero-width space = invisible bullet) as the first run, so PptxGenJS generates a bullet paragraph with `marL = bulletMarL * (1 + level)`. This aligns the continuation line's text-start position with the first line, reproducing PowerPoint Shift+Enter (soft return) behaviour.

---

### ADR-17: display:inline + borderRadius on `<code>` elements incorrectly extracted as badge shapes

**Problem**
On slides 41/42 (containing inline code) and 8 other slides, inline `<code>` elements were extracted as badge container shapes, producing unwanted solid coloured blocks in PPTX.

**Root cause**
A previous iteration (v0.1.8, commit affee6a) added `display === 'inline' && borderRadius > 0` detection to `extractInlineBadgeShapes` / `extractTextRuns`. The Marp default theme's `<code>` element has:
- `display: inline`
- `border-radius: 6px` (i.e. `0.2em` at 30px font size)
- `background: rgba(129, 139, 152, 0.12)` (alpha = 0.12)

This matched the badge detection condition. Because `rgbToHex` discards alpha, `rgba(129,139,152,0.12)` became `#818B98` (a solid grey), and PPTX received a solid grey filled block for every inline code span.

**Fix (`dom-walker.ts`)**
Added 2 guards to the `display: inline` badge branch:

1. **`inlineBorderRadius > 6` (threshold)**: `<code>` uses 6px, which is now excluded. Real badge elements (12px, 16px, 50%→14px) pass the threshold.
2. **`alpha >= 0.5` check**: Semi-transparent backgrounds like `rgba(..., 0.12)` are excluded. Real badge elements have fully opaque backgrounds.

**Tests added**
- Added: `display:inline code element (borderRadius=6, semi-transparent bg) is NOT emitted as container`
- Updated: Existing T5 tests adjusted for the `borderRadius > 6` condition.

**Affected slides (before fix)**
Slides 3, 15, 24, 32, 40, 41, 42, 50, 51, 60 (10 slides had solid grey filled blocks).

**Note**
All inline elements can legitimately have `borderRadius > 0` (e.g. `<code>`, `<abbr>`). When adding new badge detection, simultaneously add a negative test for representative "not a badge" elements (inline code). Structural negative tests are important for systematic prevention because subtle colour blocks fall within visual WARN range and are easy to miss.

*ADR-19 later superseded this threshold-based approach with a structure-first approach based on semantic tag names.*

---

### ADR-18: Treat slide numbers as a deck-wide native PowerPoint feature

**Problem**
Reconstructing Marp page numbers from `section::after` created visual drift,
duplicate numbers burned into rasterized backgrounds, and brittle behavior when
theme CSS reused `::after` for decorative bars.

**Root cause**
Slide numbers are deck metadata, not ordinary slide content. In editable PPTX,
the important behavior is automatic renumbering after reorder, copy, or delete
inside PowerPoint. Reproducing the HTML pseudo-element as fixed text overfit
the test fixture and introduced regressions.

**Fix**
- `dom-walker.ts` records only whether the source deck used pagination
  (`sourceHasPagination`)
- `slide-builder.ts` enables `slide.slideNumber` consistently for every slide
  when the source deck uses pagination at all
- `dom-walker.ts` keeps pagination pseudo-element backgrounds when they also
  provide visible decoration such as bottom bars, ribbons, or pills
- `index.ts` makes the genuine pagination text transparent before rasterizing
  any screenshot-based background so the HTML number is not burned into images
  while the decorative background itself remains visible

**Tests added**
- `pagination source detection`
- `adds PowerPoint slide numbers to every slide when the source deck uses pagination`
- `uses a fixed deck-wide placement and styling for the native slide number field`

---

### ADR-19: Structure-first inline badge detection

**Problem**
Marker-style highlights on slides 56 and 58 were extracted as rounded badge
shapes, causing visual drift and exaggerated corner radius.

**Root cause**
The previous badge detection relied too much on CSS numeric thresholds such as
`border-radius` in pixels. Semantic inline elements like `<strong>`, `<mark>`,
and `<code>` can legitimately use rounded backgrounds without representing a
separate badge/chip component.

**Fix**
- keep semantic inline tags (`strong`, `mark`, `code`) as text runs with
  `backgroundColor`
- treat only inline `span` elements as rounded inline badge candidates
- keep `inline-block` / `inline-flex` / `inline-grid` badge extraction for real
  badge components, but exclude semantic inline text tags even when CSS changes
  their display value

**Tests added**
- `display:inline strong with borderRadius:4px (slide 56/58) stays as a text highlight instead of a badge shape`
- `inline code element (borderRadius=6, semi-transparent bg) is NOT emitted as container`


---

### ADR-20: Disclose layout faithfulness limitation in public documentation

**Problem**
The README comparison table and "How it works" section implicitly claimed that
the extension faithfully reproduced HTML layout for any Marp theme. This was
contradicted by the known fact that browsers and PowerPoint use separate text
rendering engines (Chromium/Skia vs PowerPoint/DirectWrite), which produce
different character spacing and line-break calculations even for the same font.
The discrepancy is largest when a Marp theme uses web fonts (e.g. via Google
Fonts), because PowerPoint substitutes a system font, changing character widths
enough to cascade line-break shifts throughout the slide.

**Root cause (documentation)**
The limitation was known and handled in code (ADR-10: heading width expansion)
but had never been documented in user-facing documentation. End users discovering
the discrepancy would have no way to understand why it happened or what they
could do about it.

**Decision**
Added an explicit "Layout faithfulness note" paragraph in README.md, positioned
immediately after the approach explanation. The note:
- explains the separate rendering engine cause
- names web fonts as the largest risk factor
- names system fonts as a mitigation
- states that pixel-perfect match is not guaranteed

The comparison table in README.md was also updated: the Layout faithfulness row
for this extension was changed from ✅ to ⚠️ to correctly reflect the known
typographic limitation.

**Why no code fix**
The limitation is fundamental: fully reproducing PowerPoint's DirectWrite text
metrics would require access to PowerPoint's text layout engine. The only
practical mitigations (theme system fonts, heading width expansion) are already
in place. This is a documentation decision, not a code decision.

---

### ADR-21: English-only language policy — no ADR exception

**Problem**
The initial language policy read: "All source code, comments, test case names,
and documentation are written in English. Japanese may be used in the ADR log."
This exception caused ADR entries (and the section header) to be written in
Japanese. The ADR log is the most important document for understanding past
decisions; having it in a different language from the rest of the codebase
reduces its usefulness for contributors who read English.

Additionally, ADR-17 was saved in a corrupted encoding (the file was saved as
Shift-JIS in some editors), which lost the content entirely.

**Decision**
Removed the Japanese ADR exception from the language policy in:
- src/native-pptx/README.md (section header and body)
- .github/instructions/marp-editable-pptx.instructions.md
- .github/skills/marp-pptx-visual-diff/SKILL.md
- CONTRIBUTING.md (step 7)

Translated all existing ADR entries (ADR-01 through ADR-16, ADR-K01) to English.
ADR-17 was reconstructed from test names, commit context, and ADR-19 (which
superseded ADR-17's threshold-based approach).

**Why English throughout**
The project uses English for all code, comments, and test names. Mixing
languages in the ADR log creates a two-tier document where some decisions are
accessible only to Japanese readers. Encoding issues (as seen with ADR-17)
further reduce reliability of Japanese text in source files.