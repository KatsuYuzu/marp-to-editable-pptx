# Changelog

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
