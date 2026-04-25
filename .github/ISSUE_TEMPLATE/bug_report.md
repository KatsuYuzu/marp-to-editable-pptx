---
name: Bug report
about: Something doesn't look right in the exported PPTX
title: "[Bug] "
labels: bug
assignees: ''
---

## What happened?

<!-- Describe what you saw in the exported PPTX (e.g. "text on slide 3 overlaps the image") -->

## What did you expect?

<!-- Describe what the slide should look like (e.g. "the text should be below the image") -->

## Steps to reproduce

1. Open the Markdown file
2. Press F1 → **Marp: Export to Editable PPTX**
3. Open the exported `.pptx`
4. See the issue on slide N

## Environment

- OS: <!-- e.g. Windows 11 -->
- VS Code version: <!-- e.g. 1.101.0 -->
- Extension version: <!-- e.g. 1.0.0 -->
- Chrome/Edge version: <!-- e.g. Chrome 124 -->

## Marp theme

<!-- e.g. default, gaia, uncover, or custom -->

## Reproduction file

<!-- If possible, attach or paste a minimal `.md` file that reproduces the issue. -->
<!-- Please remove any confidential content — the bug is usually in the CSS/DOM structure, not the text. -->

<!-- Note: slight font differences or text wrapping at a different point than in the browser -->
<!-- are a known limitation of the export approach, not a fixable bug. See README for details. -->

```markdown
---
marp: true
---

# Paste minimal reproduction here

```
