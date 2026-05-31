# Changelog

## v1.2.0 — 2026-05-31

### New features

**Table cells with colspan/rowspan now merge correctly in PPTX**
HTML tables using `colspan` or `rowspan` attributes now produce properly merged cells in the exported PowerPoint file, matching the visual layout from the Marp preview.

**Multi-paragraph table cells now preserve paragraph breaks**
When a table cell contains multiple paragraphs (separated by `<br>` or block elements), each paragraph is now rendered as a separate text run with a line break, rather than collapsing into a single run.

**Ordered list start number (`<ol start=N>`) now reflected in PPTX**
Numbered lists beginning at a value other than 1 (e.g. `<ol start="5">`) now start at the correct number in the exported PPTX.

**List style type mapping (circle, square, lower-alpha, upper-roman, none)**
CSS `list-style-type` values are now mapped to their PowerPoint equivalents, supporting `disc`, `circle`, `square`, `none`, `decimal`, `lower-alpha`, `upper-alpha`, `lower-roman`, and `upper-roman`.

**Adjacent text elements are grouped into a single text box**
Consecutive paragraphs, lists, and blockquotes that share the same horizontal position are now merged into a single PowerPoint text box. This makes selecting and editing text in PowerPoint more natural.

**Container shapes embed their text content**
Simple card-style containers (with background/border but only plain-text children) now render as a single shape with embedded text, so the text and its background move and resize together in PowerPoint.

**Code block syntax highlighting preserved**
Syntax-highlighted code blocks now retain per-token colors, bold, and italic styling from the Marp theme in the exported PPTX.

**CSS animations and transitions are frozen before export**
Animated elements are now captured in their final state rather than mid-animation, preventing blank or partially-rendered content in the PPTX.

### Bug fixes

**Nested list bullet type now renders correctly (ul inside ol → circle)**
When an unordered list was nested inside an ordered list, items from the inner list incorrectly rendered with numbered markers. They now correctly display bullet markers (circle/disc).

**Spaces between syntax-highlighted spans no longer disappear**
In code blocks with syntax highlighting, single spaces between adjacent `<span>` elements (e.g. `async def`, `LEFT JOIN`) were silently dropped from the PPTX output. Fixed.

**Multi-line paragraph grouping no longer causes CJK text wrapping drift**
Multi-line paragraphs (where cross-engine font metrics cause different line-break positions) are now excluded from text grouping, preventing cascading Y-position errors in subsequent elements.

**Code block backgrounds inside bullet list items now render correctly**
When a code block was indented inside a bullet list item, the gray background rectangle was missing from the PPTX. Fixed.

**Bullet lists nested inside blockquotes now preserve markers and hierarchy**
List content inside blockquotes that appeared as flat plain text now exports as a proper bullet list with all levels of hierarchy preserved.

### Visual comparison improvements

The `compare-visuals` tool thresholds have been recalibrated to reflect the inherent rendering differences between Chrome (Skia) and PowerPoint (DirectWrite/GDI+):
- pixelmatch threshold: 0.12 → 0.18 (absorbs anti-aliasing and emoji color differences)
- FAIL threshold: 5% → 7% (accounts for CJK font metrics + PowerPoint COM variance)
- WARN threshold: 2.5% → 3.5%

---

## v1.1.1 — 2026-05-08

### Bug fixes

**Marker highlights (`**bold**` styled with `linear-gradient`) now appear on bold text inside bullet lists**
When a Marp theme applied a yellow-marker effect to `<strong>` using a CSS `linear-gradient` background (e.g. `strong { background: linear-gradient(transparent 62%, #fff2a8 62%) }`), the highlight was not exported to PPTX for bold text inside bullet list items. The highlight appeared correctly in paragraphs but was silently dropped from list items. Fixed.

---

## v1.1.0 — 2026-05-04

### New features

**Subscript and superscript text now renders correctly**
`H<sub>2</sub>O` and `x<sup>2</sup>` were rendered at the baseline with only a reduced font size in previous versions. They are now positioned at the correct vertical offset in the exported PPTX.

### Bug fixes

**Inline math formulas (MathJax) no longer duplicate or disappear**
When a paragraph contained both regular text and an inline math expression rendered as MathJax SVG, the surrounding text was sometimes duplicated or dropped entirely in the exported PPTX. Fixed.

**Inline code highlight restored on `![bg fit]` slides**
On slides using `![bg fit]`, all inline `<code>` elements were missing their background highlight. Fixed.

**Split-background slides now export with correct layout**
Slides using `![bg left]`, `![bg right]`, or percentage splits (`![bg right:40%]`) could have background images and text placed at incorrect positions. Fixed.

### Known limitations

On slides where text sits directly over a split background image (e.g., `![bg left]` with a colored image on the left), inline `<code>` highlights appear as a light-colored solid box rather than a semi-transparent tint. This is inherent to PowerPoint's text rendering — text run backgrounds are solid colors only. The highlight is structurally correct and is intentionally preserved.

### Various internal improvements

---

## v1.0.2 — 2026-04-26

### Bug fixes

**Emoji icons in flex layouts could overlap adjacent text**
In slides using a horizontal flex layout where an emoji icon sits beside a text element — such as a decorative icon followed by a label — the emoji's text box was extended to the right edge of the container, causing it to overlap the adjacent text in the exported PPTX. Fixed.

---

## v1.0.1 — 2026-04-25

### Bug fixes

**List item bullet markers sometimes disappeared**
List items that contained multiple inline elements — such as text followed by an emoji or bold text — could lose their bullet marker in the exported PPTX. The marker was visible in PowerPoint on Windows but invisible when the file was opened in LibreOffice. Fixed.

**Table text clipped near the right edge**
Due to font-rendering differences between Windows (DirectWrite) and Linux/macOS (Skia), table columns could clip text near the right edge. A small width margin is now applied so table text is no longer cut off.

**Border-bottom lines on custom containers missing**
`border-bottom` applied to `<div>` containers — such as custom-styled boxes or section dividers — was not rendered in the exported PPTX. It now appears correctly.

**Dashed border-bottom rendered with an opaque fill**
A heading or container styled with `border-style: dashed` on its bottom edge was drawn with a solid background fill behind the dashes, making the dashes invisible. The fill is now transparent.

**Text truncated at the right edge inside flex or grid layouts**
Text inside flex or grid child elements was occasionally cut off at the right edge. A small width slack is now applied to prevent clipping.

---

## v1.0.0 — 2026-04-11

First stable release.

This version establishes the core capability: exporting Marp Markdown presentations to fully editable PowerPoint files without requiring LibreOffice or any external office software.

### What's included

**Export to editable PPTX**
- Text boxes, images, and shapes are placed as individual native PowerPoint objects — not embedded as flat images
- Layout, fonts, colors, and positions are extracted directly from the browser's computed style, making the output theme-agnostic

**Elements supported**
- Headings, body text, and inline styling (`strong`, `em`, `code`, `mark`)
- Unordered and ordered lists, including leading badge shapes with correct alignment
- Images (raster and SVG), including images inside list items
- Tables with per-cell content
- Mermaid diagrams and other SVG content (rasterized to PNG)
- Background colors, gradient fills, and decorative shapes

**Paginated decks**
- Page numbers use PowerPoint's native slide-number field, so they renumber correctly after reordering slides
- Decorative pagination backgrounds (bars, ribbons, pills) are preserved
- Duplicate HTML page-number text nodes are suppressed

**Quality**
- 63 fixture slides with automated visual regression (pixel-diff via `compare-visuals.js`)
- 231 unit tests
- Visual comparison validated on Windows with PowerPoint COM

### Notes

- Requires Google Chrome or Microsoft Edge (no additional setup needed)
- Visual comparison in CI uses LibreOffice on Ubuntu; local comparison uses PowerPoint COM on Windows
- v1.0+ quality improvements will continue based on feedback from real-world decks
