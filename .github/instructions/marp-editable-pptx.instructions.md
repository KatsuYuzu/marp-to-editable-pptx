---
name: 'Marp Editable PPTX Development Conventions'
description: 'Development conventions for the native-pptx module of marp-to-editable-pptx. Apply when modifying dom-walker.ts or slide-builder.ts, adding tests, managing fixtures, recording ADRs, following commit conventions, or preventing regressions.'
applyTo: 'src/native-pptx/**/*.ts, src/native-pptx/test-fixtures/**, src/native-pptx/README.md'
---

# Marp Editable PPTX Development Conventions

## Language Policy

All source code, comments, test case names, documentation, and ADR entries under `src/native-pptx/` must be written in **English**. No exceptions.

Exception: test fixture content that intentionally tests Japanese character rendering (e.g., mixed Japanese/English text in `pptx-export.md`) may contain Japanese, because that text is the subject under test.
## Design Principles (Violations Are Prohibited)

**Browser is the source of truth**

- Map `getComputedStyle()` and `getBoundingClientRect()` values 1:1 to PPTX
- Do not parse Marp themes, CSS selectors, or Markdown syntax
- Element-specific hardcoding is only allowed **when the browser has already rendered the result but PPTX has a structural limitation that prevents reproduction** (e.g., SVG `<foreignObject>`, slide page numbers)
- In that case, the only permitted fix is "capture the browser rendering result as a raster image"

## Architecture

| File | Role | When to modify |
|---|---|---|
| `dom-walker.ts` | Extracts `SlideData[]` from the browser DOM | Text missing, not extracted, or extra elements mixed in |
| `slide-builder.ts` | Converts `SlideData[]` to PptxGenJS API calls | Coordinate conversion errors, PPTX output format issues, color conversion errors |
| `index.ts` | Controls the overall pipeline | Image rasterization, browser lifecycle |
| `utils.ts` | Conversion utilities (px→inch, rgb→hex, etc.) | Unit conversion errors |

> **Note**: `dom-walker.ts` is executed inside the browser via `page.evaluate()`, outside the webpack/esbuild scope.
> After any change, always run `node src/native-pptx/scripts/generate-dom-walker-script.js` to recompile.

## Build Sequence

```powershell
# Required when dom-walker.ts is changed
node src/native-pptx/scripts/generate-dom-walker-script.js

# Required when updating gen-pptx.js (local tool) after changing dom-walker.ts or index.ts
node src/native-pptx/scripts/build-native-pptx-bundle.js
```

> **`npm run build` does not run these** (it only generates the VS Code extension webpack bundle).
> Always run `generate-dom-walker-script.js` after changing `dom-walker.ts`.
> `build-native-pptx-bundle.js` generates `lib/native-pptx.cjs` which `gen-pptx.js` depends on — run it before local visual comparison.

## Fixture Management

### 🛑 Fixture Content Safety Gate — Fires When Developer's Slide Is in Context

> **Trigger**: A developer shares their slide file, pastes slide content, or describes a bug in their own deck.

At that moment, **before writing a single character of fixture text**, declare:

> "I will not reference the developer's slide content. I will write fixture text from scratch using only the approved vocabulary."

Then proceed to write the fixture using only the vocabulary in the "Compose from Scratch" rule below. Do not sanitize, paraphrase, or generalize any text from the developer's slide — even field names, status labels, numbers, or workflow step names. The CSS/HTML structure is the only thing to carry over.

### Exclude Confidential and Personal Data (Public Repository)

`src/native-pptx/test-fixtures/pptx-export.md` is committed to a public repository.
**Never include:**

- Developer local paths (`C:\Users\...`, `/home/...`)
- Customer names, project names, internal system names, or business data
- Internal URLs, IP addresses, or credentials

#### ⚠️ Text in New Reproduction Slides Must Be Composed from Scratch

**Never sanitize or generalize text copied from a developer's slide.**  
Sanitization leaves domain meaning behind and can never be exhaustive.  
Instead, **compose all text fresh using only the approved vocabulary below**, without referencing the original content at all.

**Approved vocabulary — use only these building blocks:**

| Slot type | Allowed forms |
|---|---|
| Labels / headings | `Label-A`, `Label-B`, `Col-1`, `Col-2`, `Row-N` |
| **Table column headers** | `Col-A`, `Col-B`, `Col-C`, `Col-D` … (alphabetic extension) — **never use domain terms like `N`, `Median`, `Range`, `Count`, `Total`, `Score`** |
| **Ordinal group / row labels** | `Row-1`, `Row-2`, `Item-N` — **never use `Phase N`, `Stage N`, `Step N`, `Sprint N`** |
| Category names | `Cat-A`, `Cat-B`, `Cat-C` |
| Item / task names | `Item-1`, `Item-2`, `Task-N` |
| Tag / badge text | `Tag-A`, `Tag-B`, `Tag-C` |
| Group / section names | `Group A`, `Group B` |
| Numeric values | `val-N` (e.g. `val-10`, `val-p1`) — **never use bare integers like `8`, `21`, or ranges like `(25–90)`** |
| Units / suffixes | `/uu`, `(unit)`, `(period)`, `(label-1)` |
| Sentences (short) | `Alpha beta gamma` / `Delta epsilon zeta` |
| Sentences (longer) | `Alpha item and beta gamma. Delta epsilon.` / `Zeta nu eta, theta iota kappa.` |
| Structural filler | `input`, `data`, `item`, `label`, `note`, `text` |

> **This list is closed.** Any English word not in the table above (including common verbs like `improved`, `completed`, `confirmed`, nouns like `total`, `count`, `stage`) must be replaced.

> **Scope of this rule**: Applies to all visible text inside the slide body (Markdown content, HTML element text, attribute values rendered as text). **Exempt**: slide title (`# Slide N: ...`), HTML comments (`<!-- ... -->`), and `Expected: ...` lines — these are test metadata, not slide content. Text content can be changed without affecting reproduction. If a bug does not reproduce after changing text, the cause is in the text pattern (special characters, length, line-break rules), so use a minimal reproduction text.

### Steps for Adding a Fixture

1. Confirm the issue reproduces in a standalone deck (run `gen-pptx.js` with a single-slide deck)
2. When adding `<style>`, scope it with `section` selector or similar
3. After adding to fixture, run `compare-visuals.js` for all slides to confirm existing slides are not broken

### Always Update Slide Counts in README (2 Places)

When adding slides to the fixture, always update both of the following in the same commit:

| File | Where to update |
|---|---|
| `README.md` (repository root) | The `compare-NNN.png` line and the `All slide comparisons (N slides)` count |
| `src/native-pptx/README.md` | The slide count in the "Canonical test deck" section and in the "Visual diff improvement loop" section |

> Forgetting to update the README after adding slides has happened repeatedly.
> Every slide-addition commit must include both of these updates.

## ADR Log (Required on Every Fix)

Append to the "Bug fix and decision log" section in `src/native-pptx/README.md`.
All ADR fields must be written in **English** (no exceptions, per language policy).
**Skipping the ADR and going straight to a fix will cause previously solved problems to recur — this has happened repeatedly.**

Required fields (in English):
- Problem (symptom)
- Root cause (DOM processing, CSS interpretation, coordinate calculation perspective)
- Fix (which file, function, logic was changed)
- Tests added (test case names added)
- Why it was not caught by unit tests or visual diff
## Test Conventions

- Test case names must be in **English** (per language policy)
- When fixing a bug, always add a regression test in `dom-walker.test.ts` or `slide-builder.test.ts`
- Test case names must clearly describe what is being verified
- `describe` blocks use the target function name as-is

## Two-Axis Regression Prevention

**After every fix, always regenerate the PPTX and run compare-visuals before committing.**  
"Checking compare-report.html" means running a fresh comparison against the current code — never looking at a stale report.  
This is mandatory even when the fix seems small. Visual inspection cannot be skipped.

```powershell
# 1. Rebuild bundle if dom-walker.ts or slide-builder.ts changed
node src/native-pptx/scripts/build-native-pptx-bundle.js

# 2. Regenerate HTML → PPTX → compare (run all three)
npx marp src/native-pptx/test-fixtures/pptx-export.md `
  --html --allow-local-files `
  --output src/native-pptx/test-fixtures/slides-ci.html
node src/native-pptx/tools/gen-pptx.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
node src/native-pptx/tools/compare-visuals.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
# → Open dist\compare-slides-ci\compare-report.html for visual review
```

After running compare-visuals, verify both of the following:

| Axis | What to check |
|---|---|
| ① Rule-based unit tests | Does `npx jest` pass all cases? Are previously added regression tests still passing? |
| ② Visual diff trends | In `compare-report.html`, check **the type of diff** visually. Look especially for line-break shifts, overlaps, and missing elements |

### Do Not Judge OK/NG by Diff Rate Alone

- Line-break shifts can occur with nearly 0% diff rate
- Page overflow is also not detectable by diff rate
- In visual review, always explicitly check whether the number of text lines matches the HTML

## Commit Conventions

Use Conventional Commits:

```
fix(<scope>): description
feat(<scope>): description
docs(<scope>): description
chore(<scope>): description
ci(<scope>): description
```

- scope is the target file name (e.g., `dom-walker`, `slide-builder`, `compare-visuals`)
- One commit per one problem fixed
- Do not commit files under `dist/`
- Do not commit `slides-ci.html`
- Only commit changes to `.ts` / `.test.ts` / `pptx-export.md` / `README.md`

## Branch and PR Conventions

- Branch name: `fix/description-in-kebab-case`, `feat/description-in-kebab-case`
- Merge to main via PR (no direct push)
- PR title follows the same format as commit messages
- Add a `release` label to the PR to trigger a release

## What Never to Do

- `git add` files output to `dist/`
- `git add` `slides-ci.html`
- Modify files unrelated to the fix
- Assume `npm run build` updated the bundle (after changing `dom-walker.ts`, always recompile)
- Install LibreOffice locally (use PowerPoint COM instead)
- Write element-specific processing that overrides browser CSS rendering results (violates design principles)
- Skip reading the ADR log before making a fix
- Create new tools or helper scripts without being asked (`compare-visuals.js` / `gen-pptx.js` / `diagnose-pptx.js` are sufficient)
