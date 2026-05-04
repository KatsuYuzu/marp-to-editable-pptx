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

**Structural fidelity over visual approximation**

When PPTX has a rendering limitation (e.g., solid-color-only text backgrounds, no CSS transparency):

- HTML-specified structural elements — code highlights, borders, bullets — are **always** rendered in PPTX, even when the result is visually imperfect.
- Do **not** suppress structural elements based on "the composited result looks wrong in this specific context". That is an ad-hoc visual approximation, not a structural decision.
- Known imperfections must be documented in the **"Known limitations" section of `src/native-pptx/README.md`**, not worked around with per-element special cases.
- Exception: suppress only when the element would be **entirely invisible** — i.e., its composited color is indistinguishable from the slide's visual background — AND a general rule (not a per-case heuristic) can describe it. Example: near-white (`r,g,b > 200`) highlight on a full-slide dark bg image (≥80% slide width, CSS bg is white) would be invisible against the image → suppress by that general rule (see `visualBgMayBeDark` in `slide-builder.ts`).

**Design-First Problem Solving (applies before any code change)**

When addressing a visual rendering issue, apply this checklist **before writing any code**:

1. **State the general principle**: What rule governs this entire class of rendering decisions? (e.g., "structural elements are always rendered", "DirectWrite measures fonts wider than Skia")
2. **Check if an existing principle already covers it**: If yes, apply it. If no, add the new principle to **this section** first (i.e., edit `.github/instructions/marp-editable-pptx.instructions.md` — the file you are currently reading — and include that change in the same commit as the code fix).
3. **Apply the principle consistently**: The fix must work for all cases described by the principle, not just the reported slide.
4. **Per-case heuristics are prohibited**: If the fix requires knowing the specific slide number, element bounding box relative to a specific bg image, or other per-instance data to decide whether to apply — it is an ad-hoc fix, not a design decision.

This checklist exists because the same structural mistake has recurred multiple times: responding to a reported visual issue by patching the specific case, rather than reasoning from design principles. Every fix must trace back to a general rule stated in this section or in the ADR log.

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

# Required when changing dom-walker.ts, slide-builder.ts, or index.ts
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

### Always Add New Slide Rows to README — Commit Gate

When adding slides to the fixture, always add the corresponding `<tr>` row(s) to the slide comparison table in `README.md` (repository root) **in the same commit** as the slide addition:

| File | What to update |
|---|---|
| `README.md` (repository root) | Add or update `<tr>` row(s) inside `<details>` for the new slide(s) |

**How to add rows** (the table uses 2 columns per row):
- If the current last `<tr>` has only 1 `<td>` (odd-numbered final slide), add the new slide as the second `<td>` in that existing row.
- If the current last `<tr>` already has 2 `<td>`s (even-numbered final slide), append a new `<tr>` for the next slide.
- If adding multiple slides, continue pairing: odd gets the right `<td>`, even starts a new `<tr>`.

Example — adding slide 80 when 79 is currently a solo row:
```html
<!-- Before -->
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-079.png"></td>
</tr>

<!-- After -->
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-079.png"></td>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-080.png"></td>
</tr>
```

Example — adding slide 81 when 80 is the last (even-paired) slide:
```html
<tr>
<td><img src="https://KatsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-081.png"></td>
</tr>
```

> **This is a commit gate, not a suggestion.** A commit that modifies `pptx-export.md` MUST also include the corresponding `compare-NNN.png` row addition in `README.md`. Committing these separately is prohibited.

> **Note**: `slides-ci.html` is a generated file. Do **not** commit it (see "What Never to Do"). Only commit `pptx-export.md` and `README.md`.

## ADR Log (Required on Every Fix)

Append to the "Bug fix and decision log" section in `src/native-pptx/README.md`.
All ADR fields must be written in **English** (no exceptions, per language policy).
**Skipping the ADR and going straight to a fix will cause previously solved problems to recur — this has happened repeatedly.**

Required fields (in English):
- Problem (symptom)
- Root cause (DOM processing, CSS interpretation, coordinate calculation perspective)
- Fix (which file, function, logic was changed)
- Tests added (test case names added; for fixture-only additions without `.ts` changes, write "no unit test added — fixture-only change" here)
- Why it was not caught by unit tests or visual diff

> **Scope**: ADR is required when a `.ts` file is changed (bug fixes, features). For fixture-only additions with no `.ts` change, the ADR is optional — but the "Tests added" field above must still be filled in the commit message or a brief note.
## Test Conventions

- Test case names must be in **English** (per language policy)
- When fixing a bug, always add a regression test in `dom-walker.test.ts` or `slide-builder.test.ts`
- Test case names must clearly describe what is being verified
- `describe` blocks use the target function name as-is

## Two-Axis Regression Prevention

Every fix must pass **both** axes before commit:

| Axis | What to check |
|---|---|
| ① Rule-based unit tests | `npx jest` passes all cases including previously added regression tests |
| ② Visual diff (compare-report.html) | Check the **type** of diff visually — layout shifts, overlaps, line-count mismatches, and missing elements are NG regardless of diff rate |

### Mandatory Full-Pipeline Command (run after every fix)

```powershell
# 0. If pptx-export.md was changed, regenerate the HTML fixture first
npx marp src/native-pptx/test-fixtures/pptx-export.md `
  --html --allow-local-files `
  --output src/native-pptx/test-fixtures/slides-ci.html
# (skip if only .ts files changed)

# 1. Rebuild the bundle (required if dom-walker.ts, slide-builder.ts, or index.ts changed)
node src/native-pptx/scripts/generate-dom-walker-script.js
node src/native-pptx/scripts/build-native-pptx-bundle.js
# (skip if only pptx-export.md or README changed; running them is always safe)

# 2. Run unit tests — must pass before proceeding to visual comparison
npx jest

# 3. Regenerate PPTX from the current fixture
node src/native-pptx/tools/gen-pptx.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/slides-ci.pptx

# 4. Run compare — stale images are cleaned automatically before each run
node src/native-pptx/tools/compare-visuals.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/slides-ci.pptx

# 5. Open dist/compare-slides-ci/compare-report.html and inspect visually
```

> **Why regenerate PPTX every time?** A stale PPTX (generated from a previous fixture state) produces ghost slides in the compare report — PPTX slide count diverges from HTML slide count, causing phantom MISSING entries that mask real regressions.

> **Why auto-clean?** `compare-visuals.js` cleans all `html-slide-*`, `pptx-slide-*`, `diff-slide-*` and report files at the start of each run. Never trust a report generated without a matching PPTX regeneration.

> **FAIL count is not the acceptance criterion.** Typography and anti-aliasing differences will trigger FAIL thresholds — these are **acceptable**. Pixel diff alone cannot detect line-break shifts (nearly 0% diff rate). Always verify line counts visually.

> **Human visual inspection of `compare-report.html` is the mandatory final gate.** The AI pre-check does not replace human confirmation.

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
- PR title and body must be in **English**. Provide a Japanese translation after the English body, in the same comment
- Always output PR title and body inside a code block so the user can copy them directly

**Re-output the PR title and body after every additional commit.** If a PR title and body have already been shown and a new commit is added to the branch, output a fresh PR title and body that reflects the full current state of the branch. Do not assume the user still has the previous output visible.

## Changelog Conventions

`CHANGELOG.md` is read by **end users of the VS Code extension**, not by developers.

**Describe the visible symptom — what the user saw or could not do:**

| ✅ Write this | ❌ Not this |
|---|---|
| "List item bullet markers sometimes disappeared when the list contained inline elements such as bold text or emoji" | "Fixed ADR-29: propagate `bullet`/`indentLevel` to all runs in `toListTextProps` in `slide-builder.ts`" |
| "Text inside flex child elements was occasionally truncated at the right edge" | "Added slack in `dom-walker.ts` for flex/grid children (`emojiWidthOverride`)" |

**Internal details belong in the ADR log, not the changelog:**
- File names (`dom-walker.ts`, `slide-builder.ts`), function names, ADR numbers, OOXML details → ADR only
- Multiple small internal fixes that have no user-visible symptom → summarize as "Various internal fixes and improvements" — do not enumerate them

**Format for each entry:**
- **Heading**: short symptom phrase (plain English, no period)
- **Body**: 1–2 sentences. What the user experienced. What is now correct. End with "Fixed." for bug fixes or "Improved." for improvements.

## What Never to Do

- `git add` files output to `dist/`
- `git add` `slides-ci.html`
- Commit `pptx-export.md` without simultaneously adding the corresponding `compare-NNN.png` row(s) to `README.md` — partial commits that leave the comparison table out of sync are prohibited
- Modify files unrelated to the fix
- Assume `npm run build` updated the bundle (after changing `dom-walker.ts`, `slide-builder.ts`, or `index.ts`, always recompile)
- Install LibreOffice locally (use PowerPoint COM instead)
- Write element-specific processing that overrides browser CSS rendering results (violates design principles)
- Skip reading the ADR log before making a fix
- Create new tools or helper scripts without being asked (`compare-visuals.js` / `gen-pptx.js` / `diagnose-pptx.js` are sufficient)
- Assume local PowerPoint COM comparison catches all bugs — OOXML structural issues (e.g., duplicate `<a:pPr>` in bullet runs) may pass locally but fail in LibreOffice CI (see ADR-29)

## Specifications (see `src/native-pptx/README.md`)

The following are defined in the README under "Known limitations":
- **Unsupported Marp features** (e.g., `paginate:hold`) — do not use in test fixtures
- **Marp Markdown pitfalls** (e.g., `***` = slide separator) — use `<hr>` instead
- **Compare tool limitations** (pagination key dedup, pixel diff blind spots)

Read these before adding fixture slides or interpreting compare results.

## Agent Workflow: Auto-actions After Fix

The agent must perform these actions **automatically after every code change** without waiting for user instruction:

### 0. Create a branch before any work starts (auto, no user instruction needed)

**Do not wait for the user to say "branch", "checkout", or any git instruction.** Branch creation is step zero of every task.

```powershell
git checkout -b fix/<description>   # for bug fixes
git checkout -b feat/<description>  # for feature or fixture additions
```

Naming: `fix/description-in-kebab-case` or `feat/description-in-kebab-case`. Do not start any code or file change before this step.

### 1. Regenerate compare report and present results

After any change to `dom-walker.ts`, `slide-builder.ts`, `index.ts`, or `pptx-export.md`:

```powershell
# Steps 1-2 are only strictly needed for .ts changes, but running them always is safe (no-op if unchanged)
node src/native-pptx/scripts/generate-dom-walker-script.js
node src/native-pptx/scripts/build-native-pptx-bundle.js
npx marp src/native-pptx/test-fixtures/pptx-export.md `
  --html --allow-local-files `
  --output src/native-pptx/test-fixtures/slides-ci.html
node src/native-pptx/tools/gen-pptx.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
node src/native-pptx/tools/compare-visuals.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
```

**Always report to user:**
- Full path to the report: `dist\compare-slides-ci\compare-report.html`
- Summary line: `FAIL N, WARN N, OK N, MISSING N`
- Per-slide diff% changes for any slide that was targeted by the fix (before → after)
- Any unexpected changes in slides that were NOT targeted

> **Do not declare a task complete without running compare-visuals and reporting the summary line.** "I made the change" is not done. "I ran compare: FAIL 0, WARN N, OK N, MISSING 0" is done.

### 2. Track per-slide diff% before and after

Before starting a fix, run compare-visuals and record the diff% for affected slides.
After the fix, regenerate and compare. Report format:

```
| Slide | Before | After | Status |
|-------|--------|-------|--------|
| 75    | FAIL 41.64% | WARN 1.64% | ✓ fixed |
| 77    | FAIL 44.37% | WARN 1.38% | ✓ fixed |
```

> **If no prior compare run exists in this session:** note "baseline unavailable" and report only the current "after" values. Do not attempt to reconstruct baseline from git history (`dist/` is gitignored).

### 3. Git commit without being asked

After tests pass AND compare report shows no new regressions:
1. Add a regression test in `dom-walker.test.ts` or `slide-builder.test.ts` (English test name, describing what is verified). **Exception**: if no `.ts` file was changed (e.g., fixture-only fix), a regression test is not required — note the reason in the ADR "Tests added" field instead.
2. Stage only the source files (never `dist/`, `slides-ci.html`, or generated build outputs)
3. Commit with a conventional commit message describing the fix
4. Report the commit hash to the user

**Do not wait for user to say "commit" or "git"** — committing is part of completing a fix.

> **Commit vs human confirmation**: The agent commits source code automatically. However, `compare-report.html` is always presented to the user as an FYI — if the user spots a visual problem after commit, `git revert` is used. The commit is **not** blocked on explicit user approval; the human gate is an async quality check, not a synchronous blocker.

### 4. ADR and CHANGELOG (auto-include in same commit)

When a bug fix is committed, the same commit must also include:
- **ADR entry** in `src/native-pptx/README.md` (English, required fields)
- **CHANGELOG entry** in `CHANGELOG.md` (user-visible symptom, English)

These are not separate follow-up actions — they ship with the code change.

### 5. `--html` flag is mandatory

The test fixture uses `html: true` in frontmatter (for `<hr>`, `<style scoped>`, etc.). The marp CLI must always be invoked with `--html`. Without it, HTML elements render as literal text and the compare will show false FAILs.
