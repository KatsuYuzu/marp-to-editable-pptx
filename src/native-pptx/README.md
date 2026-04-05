# native-pptx — Editable PPTX without LibreOffice

This module generates fully editable PowerPoint files (`.pptx`) from Marp HTML
output using browser-DOM extraction + PptxGenJS, with no dependency on
LibreOffice or any external office converter.

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
  6. rasterizeSlideTargets()       screenshot CSS-filtered images
  7. buildPptx(slides)  ──►  PptxGenJS buffer
  └──► .pptx file
```

### File map

| File                             | Role                                                                                                                         |
| -------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `index.ts`                       | Pipeline orchestration: browser launch → DOM extract → rasterize → build PPTX                                                |
| `dom-walker.ts`                  | `extractSlides()` — runs **in the browser via `page.evaluate()`**; reads DOM, returns `SlideData[]`                          |
| `dom-walker-script.generated.ts` | Compiled IIFE string of `dom-walker.ts` — injected via `addScriptTag` (regenerate with `npm run generate:dom-walker-script`) |
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
> `npm run generate:dom-walker-script` to update the generated file.

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

### Background screenshots (CSS filter rasterization)

Slide backgrounds that use CSS filters (`grayscale`, `brightness`, `blur`, etc.)
or complex CSS that pptxgenjs cannot reproduce natively are captured via
Puppeteer screenshot and embedded as images. Three rasterization passes run in
`index.ts`:

1. `buildFilteredBgJobs` — `<figure>` background images with CSS filters
2. `buildCssFallbackBgJobs` — CSS `background-image` set by Marp directives
3. `buildFilteredContentImageJobs` — inline `<img>` with CSS filters

Before screenshots are taken, the bespoke **OSC overlay**
(`<div class="bespoke-marp-osc">`) and note panels (`.bespoke-marp-note`) are
hidden via `page.addStyleTag()` so they do not appear in the output.

### Mermaid diagrams

marp-cli に `--html` を渡した場合、Markdown 上の `<div class="mermaid">` + `<script src="mermaid.min.js">` がそのまま HTML に出力される。Puppeteer が `waitUntil: 'networkidle0'` で CDN スクリプトの読み込みを待ち、さらに 1 秒の settle 待機で mermaid の非同期レンダリング（Promise ベース）が完了する。その後、DOM ウォーカーが `<svg>` を data URL に変換してキャプチャする。

mermaid@10 は flowchart テキストラベルを `<foreignObject>` 経由で描画する。PowerPoint は `<foreignObject>` を SVG としてネイティブ描画できないため、DOM ウォーカーは `<svg>` 内に `<foreignObject>` が存在する場合に `rasterize: true` を設定し、`index.ts` がその要素を PNG スクリーンショットに置き換える。

**前提条件**：拡張の設定 `markdown.marp.html` が `"all"` のとき自動的に `--html` が marp-cli に渡される（`extension.ts` で制御）。`"all"` 以外の場合、スクリプトタグはセキュリティポリシーにより marp-cli が除去するため mermaid は動作しない。これは意図した動作である（marp for VS Code と同じ挙動）。

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
this module. It contains 59 slides covering every known rendering challenge:

- Basic headings, paragraphs, lists, tables, code blocks
- `border-bottom` on H1 and `border-left` vertical bar on H2/H3
- Inline badge shapes (pill, circle, status chips, step numbers)
- Leading-badge heading offset (`computeLeadingOffset`)
- ZWJ emoji sequences (🧑‍💻, 👨‍👩‍👧‍👦 — multi-codepoint single glyph)
- `strong { background-color: ... }` solid colour highlight
- CSS `section::before/after` banner suppression
- Complex background filters (`blur`, `brightness`, `grayscale`)
- Split-layout (`flex`, `grid`, HTML `<div>`) with images

Use this file as the input when running the visual diff loop below. The
file lives inside the repository so any contributor can reproduce results.

### Prerequisites

1. Install dependencies: `npm install`
2. Build the native-pptx bundle: `npm run build:native-pptx`
3. Have Chrome/Chromium available (auto-detected via `@puppeteer/browsers`)
4. Install LibreOffice for PPTX → PNG rendering, **or** use PowerPoint COM
   automation on Windows

### Steps

```sh
# 1. Convert the canonical test deck to bespoke HTML
npx marp src/native-pptx/test-fixtures/pptx-export.md --html --output /tmp/pptx-export.html

# 2. Generate native PPTX from the HTML
node src/native-pptx/tools/gen-pptx.js /tmp/pptx-export.html /tmp/pptx-export.pptx

# 3. Compare HTML slides vs PPTX slides side-by-side
#    (requires LibreOffice or PowerPoint COM for PPTX screenshots)
node src/native-pptx/tools/compare-visuals.js /tmp/pptx-export.html /tmp/pptx-export.pptx

# Output in /tmp/compare-pptx-export/:
#   html-slide-000.png   — Marp HTML reference screenshot
#   pptx-slide-000.png   — PPTX slide screenshot
#   compare-000.png      — side-by-side diff image
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

## バグ修正・意思決定の記録 (ADR log)

このセクションはバイブコーディングの再現性を確保するための意思決定ログ。
「なぜそうなっているか」を理解せずに同じコードを書き直すと同じバグを再度生む可能性があるため、
発見された問題とその根本原因・決定を記録する。

### ADR-01: toListTextProps が highlight を出力しなかった（slide 56/58）

**問題**  
`<strong style="background-color:#f1c40f">` が `<li>` 内にある場合、PPTX のリスト項目テキストにハイライト色が適用されなかった。

**根本原因**  
`slide-builder.ts` の `toListTextProps` で `run.backgroundColor` を `highlight` に変換する処理が欠落していた。`toTextProps`（段落用）では正しく実装されていたが開発時に `toListTextProps` への対応が漏れた。

**DOM walker 側は正常だった**  
`extractListItems` が `extractTextRuns(<strong>)` を呼ぶと、`extractTextRuns` は要素自身の `elementBg`（`rgb(241, 196, 15)`）を TEXT_NODE の `bg` として伝播するため、`run.backgroundColor` は正しく設定されていた。問題は builder 側にのみ存在した。

**修正**  
`toListTextProps` の run マッピングに `highlight: run.backgroundColor ? rgbToHex(run.backgroundColor) : undefined` を追加。

**なぜ検知されなかったか**  
`toListTextProps` のテストに `backgroundColor` を持つ run が含まれていなかった。`toTextProps` の `highlight` テストは存在したが `toListTextProps` には同等のテストがなかった。

**再発防止**  
テスト `toListTextProps — backgroundColor の run には highlight が設定される` を追加。これにより build 時点でリグレッションを検知できる。

---

### ADR-02: mermaid SVG の foreignObject が PowerPoint で描画されなかった

**問題**  
mermaid@10 の flowchart が PPTX で表示されるが、テキストラベルと矢印が描画されない。

**根本原因**  
mermaid@10 は flowchart ノードのテキストを `<foreignObject><div>...</div></foreignObject>` で描画する。PowerPoint は `<foreignObject>` を含む SVG を埋め込むと foreignObject 部分を無視するため、テキストが消える。

DOM walker の `tag === 'svg'` ブランチは SVG 全体を base64 エンコードして data URL に変換するが、`rasterize: true` を設定していなかった。一方 `tag === 'pre'` ブランチ（コードフェンス内 mermaid）は既に `rasterize: true` を設定しており不整合があった。

**修正**  
`tag === 'svg'` ブランチで `child.querySelector('foreignObject') !== null` を確認し、true なら `rasterize: true` を設定。`index.ts` の `buildRasterizeImageJobs` がこれを Puppeteer スクリーンショット（PNG）に置き換える。

**なぜ検知されなかったか**  
CI の比較画像（compare-NNN.png）でのみ外見上確認できる問題だった。単体テストには SVG+foreignObject のケースが存在せず、CI の visual diff にも閾値判定がなかった。

**再発防止**  
1. CI に per-slide RMSE レポートを追加（RMSE > 0.20 で警告出力）
2. `dom-walker.test.ts` に `<svg>` + `<foreignObject>` が `rasterize: true` になるテストを追加することが望ましい（TODO）

---

### ADR-03: mermaid を div.mermaid 記法で書くべき、コードフェンスではない

**問題**  
test fixture で mermaid を ` ```mermaid ` コードフェンスで書いたが PPTX に変換されなかった。

**根本原因**  
marp-core は ` ```mermaid ` を出力する mermaid 変換プラグインを持たない（marp-core 4.x）。コードフェンスは `<pre><code class="language-mermaid">` として出力されるだけで SVG にはならない。

**正しい方法**  
`<div class="mermaid">diagram code</div>` が正しい記法。mermaid.js（CDN から読み込み）がこの div を SVG に変換する。スクリプトタグは同じスライドか先行するスライドに一度だけ置けばよい。

**前提条件**  
frontmatter に `html: true`、かつ VS Code 設定 `markdown.marp.html: "all"` が必要。`html: true` だけでは `<script>` タグを marp-cli がストリップする（セキュリティポリシー）。

---

### ADR-04: 一時 HTML ファイルをソース MD と同じディレクトリに出力する

**問題**  
相対パスの画像（`./attachments/image.png` など）が PPTX に反映されなかった。

**根本原因**  
`os.tmpdir()` に一時 HTML を出力すると、marp-cli が画像の相対パスを一時ディレクトリ基準で解決してしまい、元の MD ファイルの隣にある画像が見つからなかった。

**修正**  
一時 HTML を `path.dirname(doc.uri.fsPath)` に出力するよう変更。拡張後は cleanup で削除。

---

### ADR-05: CI の視覚的回帰テスト設計

**現状**  
CI（`screenshots.yml`）は HTML スクリーンショットと PPTX スクリーンショットの横並び Compare 画像を生成し gh-pages に公開する。人間が目視確認する運用。

**問題**  
目視確認では漏れが発生する（実際に highlight バグが長期間気付かれなかった）。

**決定した改善方針**  
1. **per-slide RMSE レポート**：CI に ImageMagick `compare -metric RMSE` ステップを追加。RMSE > 0.20 のスライドを WARN として CI ログに出力。このステップは `continue-on-error: true` で job を落とさない（HTML vs PPTX は常に差が生じるため）
2. **単体テストで構造を保証**：視覚的な問題の根本は `SlideData[]` の構造問題（`backgroundColor` の欠落など）にあることが多い。単体テストで `SlideData[]` の構造を厳密に検証することで、ブラウザ不要で CI 前に検知できる
3. **test fixture 59 slides**：`pptx-export.md` はすべての既知エッジケースを網羅した canonical fixture。新しいバグを発見したら必ず対応するスライドをこのファイルに追加してから修正する

**将来的な検討（TODO）**  
- PPTX-to-PPTX の regression（同コミット間比較）を追加し、純粋な PPTX レンダリングの後退を検知する
- `pixelmatch` による差分画像生成（赤で差分ハイライト）を追加する
