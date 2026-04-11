# Changelog

## 1.0.0 - 2026-04-11

Initial stable release.

### Highlights

- Export Marp Markdown decks to editable PowerPoint slides with native text, image, and shape placement.
- Cover 63 fixture slides with the local visual-diff workflow and regression tests.
- Use PowerPoint's native slide-number field for paginated decks so page numbers renumber correctly after reordering.
- Preserve decorative pagination backgrounds such as bars, ribbons, and pills while suppressing duplicate HTML page-number text.
- Keep semantic inline highlights such as `strong`, `mark`, and `code` as text highlights instead of detached badge shapes.
- Preserve leading badge spacing in lists so extracted badge shapes do not overlap list text.

### Validation

- Unit tests: 231 passing
- Local build: passing
- Local visual comparison: 63 slides compared, 0 FAIL

### Notes

- Local visual comparison is validated on Windows with PowerPoint COM.
- Post-1.0.0 quality improvements will continue based on user feedback from real decks.