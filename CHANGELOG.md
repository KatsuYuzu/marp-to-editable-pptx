# Changelog

## v1.0.1 — 2026-04-25

### Bug fixes

**List item bullet markers sometimes disappeared**
List items that contained multiple inline elements (e.g., text followed by an emoji or bold text) could lose their bullet marker in the exported PPTX. The marker was visible in PowerPoint on Windows but invisible when the file was opened in LibreOffice. Fixed.

**Table column widths slightly too narrow**
Due to font-rendering differences between Windows (DirectWrite) and Linux/macOS (Skia), table columns could clip text near the right edge. A small width margin is now added so table text is no longer cut off.

**Border-bottom lines on custom containers missing**
`border-bottom` applied to `<div>` containers — such as custom-styled boxes or section dividers — was not rendered in the exported PPTX. It now appears correctly.

**Dashed border-bottom rendered with an opaque fill**
A heading or container with `border-style: dashed` on its bottom border was drawn with a solid background fill behind the dashes, making the dashes invisible. The fill is now transparent.

**Text clipped in flex/grid layouts**
Text inside flex or grid child elements was occasionally truncated at the right edge. A small width slack is now applied to prevent clipping.

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
