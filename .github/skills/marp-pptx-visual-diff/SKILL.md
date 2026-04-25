---
name: marp-pptx-visual-diff
description: 'Skill for running the visual fidelity improvement loop for marp-to-editable-pptx on Windows. No LibreOffice required. Uses PowerPoint COM for PPTX→PNG conversion. Covers design principles (browser as source of truth), language policy (English only), fix location decisions (dom-walker vs slide-builder), fixture confidential data exclusion, README 2-place updates, bundle regeneration, compare-visuals visual review, and ADR recording — all in a single workflow. Use when: "PPTX output looks wrong", "slide layout is shifted", "text is missing", "want to run compare-visuals", "visual diff loop", "line breaks are shifting", "text wrapping differs from browser", "diff rate alone is insufficient to judge". Never install LibreOffice. Never include confidential or personal data in fixtures.'
argument-hint: 'Symptom description (e.g., "Slide 34 has a floating badge") or target slide number'
---

# marp-pptx Visual Diff Improvement Loop (Windows / No LibreOffice)

## Environment Prerequisites

- **Windows-only workflow**. PPTX → PNG conversion uses PowerPoint COM.
- **Never install LibreOffice**. The raison d'être of this repository is "editable PPTX without LibreOffice" — the environment assumes this as well.
- CI runs on Ubuntu + LibreOffice. Local verification uses PowerPoint COM as a substitute; exact numerical parity with CI is not expected. CI-side `compare-061.png` etc. are obtained by manually triggering GitHub Actions.

## Design Principles (Verify Before Starting Any Fix)

**Browser is the source of truth**

- Map `getComputedStyle()` / `getBoundingClientRect()` values 1:1 to PPTX
- Do not parse themes, CSS selectors, or Markdown syntax
- Element-specific hardcoding is only allowed when "the browser has already rendered the result but PPTX has a structural limitation that prevents reproduction"
  - Allowed examples: SVG `<foreignObject>` (PowerPoint cannot render), slide page numbers (re-numbering required)
  - The only permitted fix in that case is "capture the browser rendering result as a PNG"
- Fixes that violate this principle (code that interprets CSS, element-specific branching) are architectural errors

## Fix Location Decision (dom-walker vs slide-builder)

| Symptom | Fix location |
|---|---|
| Text not extracted, missing, or extra elements mixed in | `dom-walker.ts` |
| **Text appears twice (duplicated)** — same text specified 2+ times in PPTX | `dom-walker.ts` (likely mis-collecting pre-render text nodes) |
| Coordinate conversion errors, width/height calculation errors | `dom-walker.ts` or `slide-builder.ts` |
| PPTX output format issues (margins, colors, fonts) | `slide-builder.ts` |
| Image rasterization condition missing | `index.ts` |

> `dom-walker.ts` runs inside the browser. After any change, always re-run `generate-dom-walker-script.js`.

## Language Policy

All source code, comments, test case names, and ADR entries under `src/native-pptx/` must be written in **English**. No exceptions.

Exception: test fixture content that intentionally tests Japanese character rendering (e.g., mixed Japanese/English text in `pptx-export.md`) may contain Japanese, because that text is the subject under test.
## Overall Loop Flow

```
Confirm symptom
  │
  ├─ Run `npx jest` first to assess current state
  │    │
  │    ├─ Tests failing → minimal reproduction in dom-walker.test.ts / slide-builder.test.ts → fix → test → rebuild bundle → compare
  │    │
  │    └─ Tests passing (visual issue only)
  │         │
  │         ├─ Step 1: Add reproduction slide to fixture (exclude confidential data, update README 2 places)
  │         ├─ Step 2: First compare runs with existing bundle. Rebuild only after modifying dom-walker.ts
  │         ├─ Step 3: Generate HTML (--html --allow-local-files required)
  │         ├─ Step 4: Generate PPTX (gen-pptx.js)
  │         ├─ Step 5: Compare with compare-visuals
  │         ├─ Step 5b: Visually classify diff type (do not judge by diff rate alone; watch for line-break shifts)
  │         ├─ Step 5c: Check ADR → fix → 2-axis regression check (tests + visual)
  │         ├─ Step 6: Fix (dom-walker.ts or slide-builder.ts with English regression tests)
  │         └─ Step 7: Record ADR
  │
  └─ compare passes but user reports an issue
       │
       ├─ Step 5b visual check → no issue → ask user to provide a screenshot of the PPTX in PowerPoint
       └─ Step 5b visual NG → Step 5c (check ADR) → Step 1 (add fixture) → standard loop from Step 3
```

---

## Step 1: Add Reproduction Slide to Fixture

**When you discover a new bug, always add a fixture first. Even when a bug is found in an existing slide, add a new minimal reproduction slide at the end — do not modify existing slides directly.**

```
src/native-pptx/test-fixtures/pptx-export.md
```

- Add after the last existing slide, separated by `---`.

### 🛑 Fixture Content Safety Gate — Check Before Writing Anything

> **If the developer's slide content is visible in the current context (pasted text, open file, screenshot), stop here.**

Never copy, adapt, or sanitize that content. Sanitization cannot remove domain meaning and can never be exhaustive.

**Mandatory action**: Close the mental reference to the developer's slide. Then write ALL fixture text using only the approved vocabulary below — without looking at the original content.

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

> **Scope of this rule**: Applies to all visible text inside the slide body. **Exempt**: slide title (`# Slide N: ...`), HTML comments (`<!-- ... -->`), and `Expected: ...` lines — these are test metadata, not slide content.

The only information to carry over from the developer's slide is:
- The CSS/HTML **structure** (class names, nesting, layout properties)
- The **text length category** if the bug is length-triggered (short / medium / long)
- **Special characters** if the bug is character-triggered (e.g., emoji, ZWJ, CJK)

Text meaning is never relevant to reproducing a layout bug.
- Include the slide number and bug description in the slide title (e.g., `# Slide 62: ...`).

### ⚠️ Always Update README in 2 Places (a Repeated Failure Pattern)

Forgetting to update the README after adding slides has happened repeatedly. Update both of the following in the same commit:

| File | What to update |
|---|---|
| `README.md` (repository root) | Add `compare-NNN.png` line inside `<details>` and update `All slide comparisons (N slides)` count |
| `src/native-pptx/README.md` | Slide count in "Canonical test deck" section (e.g., `63 slides`) and fixture file description |

The CI `screenshots.yml` auto-updates comparison images on GitHub Pages, so just adding the `<img>` tag for `compare-NNN.png` is sufficient — CI generates the actual image.

### Mandatory Rules for Fixture Content

#### Exclude Confidential and Personal Data (Public Repository)

This file is committed to a public repository. **Never include:**

- Developer local paths (`C:\Users\...`, `/home/...`, etc.)
- Customer names, project names, internal system names
- Business data or real data (file names, amounts, names, IDs, etc.)
- Internal URLs, IP addresses, or credentials

#### ⚠️ [critical] Text in New Reproduction Slides Must Be Composed from Scratch

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

**Rules:**

1. Open a blank text editor mentally. Write text using only the vocabulary above.
2. Do **not** look at the developer's slide content while writing fixture text.
3. Numeric values must look obviously synthetic (`val-50`, `val-1000`) — never use realistic-looking numbers, currency symbols, or percentages.
4. If the bug is triggered by text length or pattern (e.g. long word, special char, line-break), reproduce that property using the approved vocabulary instead of the original text.

The root cause of bugs lies in CSS layout and DOM structure. Text content can be changed without affecting reproduction. If a bug does not reproduce after changing text, the cause is in the text pattern (special characters, length, line-break rules), so use a minimal reproduction text.

#### Scoping (Slide-Specific vs Global CSS)

Before adding a fixture, confirm:

1. **Check whether the bug reproduces in a standalone 1-slide Markdown deck**
   ```powershell
   # Create dist/repro-single.md and verify with a standalone build
   npx marp dist/repro-single.md --html --allow-local-files --output dist/repro-single.html
   node src/native-pptx/tools/gen-pptx.js dist/repro-single.html dist/repro-single.pptx
   ```
   - Reproduces -> slide-specific issue -> add only that CSS/DOM structure to fixture
   - Does not reproduce -> another slide's `<style>` or global CSS is interfering -> identify the source before designing the fixture

2. **Scope `<style>` to the slide page**  
   A `<style>` block that modifies theme or global CSS will change the appearance of all subsequent slides.
   - When adding `<style>`, always scope it with `section` selector or similar
   - Or use Marp's `<!-- _class: xxx -->` to apply only to that slide

3. **After adding the fixture, run compare for all slides to confirm existing slides are not broken**
   ```powershell
   node src/native-pptx/tools/compare-visuals.js `
     src/native-pptx/test-fixtures/slides-ci.html `
     dist/compare-out.pptx
   # -> Check compare-report.html for any new FAILs
   ```
   If new FAILs appear, suspect global impact and review the added `<style>` or HTML structure.
---

## Step 2: Required Builds

> **The initial symptom comparison (Steps 3–5) can run with the existing bundle.**
> Bundle regeneration is only needed after modifying `dom-walker.ts`.

```powershell
# When dom-walker.ts has been changed
node src/native-pptx/scripts/generate-dom-walker-script.js

# When regenerating the standalone bundle used by gen-pptx.js
# ← npm run build does NOT do this (a common trap)
node src/native-pptx/scripts/build-native-pptx-bundle.js
```

> **Note**: `npm run build` generates the VS Code extension webpack bundle, but does NOT regenerate `lib/native-pptx.cjs` (the bundle that `gen-pptx.js` depends on). Always run the two commands above after changing `dom-walker.ts`.

---

## Step 3: Generate HTML

```powershell
# --html and --allow-local-files are required (omitting them breaks badges, mermaid, images)
npx marp src/native-pptx/test-fixtures/pptx-export.md `
  --html --allow-local-files `
  --output src/native-pptx/test-fixtures/slides-ci.html
```

`slides-ci.html` is in `.gitignore`. Do not commit it.

---

## Step 4: Generate PPTX

```powershell
node src/native-pptx/tools/gen-pptx.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
```

When narrowing down the cause, first verify with a 1-slide deck:

```powershell
# Create a temporary markdown with just the problem slide
npx marp dist/slide-repro.md --html --allow-local-files --output dist/slide-repro.html
node src/native-pptx/tools/gen-pptx.js dist/slide-repro.html dist/slide-repro.pptx
```

#### Debug Output (DOM Extraction JSON)

```powershell
$env:MARP_PPTX_DEBUG = '1'
node src/native-pptx/tools/gen-pptx.js dist/slide-repro.html dist/slide-repro.pptx
# → Dumps SlideData[] to dist/slide-repro.native-pptx.json
```

Verify that coordinates and text are correctly extracted from the JSON before writing tests — this speeds up root cause identification.

---

## Step 5: Compare with compare-visuals

```powershell
# PPTX → PNG uses PowerPoint COM (requires PowerPoint to be installed)
node src/native-pptx/tools/compare-visuals.js `
  src/native-pptx/test-fixtures/slides-ci.html `
  dist/compare-out.pptx
```

Output always goes to `dist/compare-slides-ci/` (never written to `src/`):

| File | Contents |
|---|---|
| `html-slide-NNN.png` | Reference screenshot of Marp HTML |
| `pptx-slide-NNN.png` | PowerPoint COM output |
| `diff-slide-NNN.png` | pixelmatch diff |
| `compare-report.html` | Per-slide diff rate summary |

Open `compare-report.html` and check FAIL / WARN slides.

---

## Step 5b: Classify the Type of Diff (Never Judge OK/NG by Diff Rate Alone)

**Diff rate (RMSE, pixel diff%) is only a reference value.** Whether the result is acceptable as a presentation slide depends on "what is shifted and how" -- not the number.

### Acceptable Diffs (No Fix Required)

| Type | How to identify |
|---|---|
| Font anti-aliasing difference | Only character outlines are red in diff-NNN.png. Layout matches |
| Sub-pixel level position shift | 1-2px uniform blur. Elements are in the correct position overall |
| Inter-OS/browser kerning difference | Slightly different character spacing but still within the line |
| Background gradient minor difference | Diff is uniformly faint and spread out. Not at shape or character boundaries |

### Diffs That Require Fixing (NG)

| Type | How to identify | Typical root cause |
|---|---|---|
| **Layout position shift** | Element is entirely shifted horizontally/vertically within the slide. "Band of red" in diff | Coordinate calculation error, padding/margin not accounted for |
| **Shape/text overlap** | Elements that should not overlap are overlapping | z-order, double-counted coordinate offset |
| **Line wrap / line count mismatch** | PPTX text overflows or has more/fewer line wraps. Vertical red line at line end in diff | Text box width/height insufficient, font size conversion error |
| **Extra text element mixed in** | Text present in PPTX but not in HTML (e.g., raw source code overlaid) | DOM walker mis-collecting pre-render text nodes |
| **Shape/image missing** | Element present in HTML is gone in PPTX. Solid red block in diff | Incorrect skip condition in extract logic |
| **Shape color/fill mismatch** | Background or border color is significantly different. Large strong red area in diff | backgroundColor retrieval error, incorrect transparency handling |

### Classification Procedure

1. List slides with high diff rate in `compare-report.html`
2. **Visually compare** `html-slide-NNN.png` and `pptx-slide-NNN.png` side by side for each slide
3. Classify the pattern in `diff-slide-NNN.png` -- "where is the diff occurring"
4. **Also visually check slides with low diff rate (not FAIL)**
   - NG-class problems can be hidden even when the diff rate is low
   - Text overlap and missing shapes in particular may not appear in diff rate
5. Record NG-classified issues as fix targets before starting to fix

### "Low Diff Rate = OK" Is Prohibited

- RMSE shows the "quantity" of diff, not the "type"
- One completely missing shape can result in a low diff rate if everything else matches
- Conversely, only font rendering differences can trigger a FAIL
- **Review every slide visually before declaring "this slide passes"**

> If a problem cannot be identified by visual inspection, ask the user to provide a screenshot of the PPTX opened in PowerPoint, and identify where the PPTX display differs from the HTML.

### Warning: Line-Break / Wrap Shifts Can Have Nearly 0% Diff Rate

This is the most critical item that has been repeatedly missed.

- A single line-break increase/decrease in text causes all subsequent elements to shift vertically
- Wrap shifts **can occur with nearly 0% diff rate** (no diff appears when surrounding pixels are the same color)
- Page overflow is also undetectable by diff rate
- During visual review, explicitly verify "does the number of text lines match the HTML" and "is there any text outside the bounding box"

**Visual Checklist (verify regardless of diff rate):**
- [ ] Number of text lines matches the HTML for each slide
- [ ] No text overflow outside text boxes
- [ ] Continued lines of bullet points do not overlap with the next element
- [ ] Inline elements such as emoji and badges stay on the same line
---

## Step 5c: Two-Axis Regression Check

**Check the ADR before fixing. After fixing, verify no regressions on both axes.**

| Axis | What to check |
|---|---|
| ① Rule-based unit tests | Do `dom-walker.test.ts` / `slide-builder.test.ts` all pass? Are previously added regression tests still passing? |
| ② Visual diff trends | In `compare-report.html`, check **the type of diff** (not the rate). Visually verify especially line-break shifts, overlaps, and missing elements |

### Required Checks Before Fixing

1. Read the ADR log in `src/native-pptx/README.md` to understand past decisions and already-resolved cases
2. Confirm the proposed fix does not contradict existing ADR decisions (if it does, supersede with a new ADR rather than reverting)

> **Skipping the ADR log before fixing leads to repeated regressions.** When a previously solved problem recurs, it is usually because the fix was not recorded in an ADR.

---

## Step 6: Fix Strategy

| Problem type | Where to fix |
|---|---|
| Text missing / shifted / extra elements mixed in | `src/native-pptx/dom-walker.ts` |
| Coordinate conversion, width/height calculation errors | `src/native-pptx/dom-walker.ts` or `src/native-pptx/slide-builder.ts` |
| PPTX output format issues (colors, margins, fonts) | `src/native-pptx/slide-builder.ts` |
| Image rasterization condition missing | `src/native-pptx/index.ts` |

### Steps After Fixing

1. Add a **regression test in English** to `dom-walker.test.ts` or `slide-builder.test.ts`
2. If `dom-walker.ts` was changed: run `node src/native-pptx/scripts/generate-dom-walker-script.js`
3. Re-run Step 2 -> Step 4 -> Step 5 and verify the diff has improved
4. Append an ADR to the ADR log in `src/native-pptx/README.md`

### What to Include in the ADR

- Problem (symptom) and root cause (from the perspective of DOM processing, CSS interpretation, coordinate calculation)
- Fix (which file, function, logic was changed)
- Tests added (test case names added)
- **Why the unit tests or visual diff did not catch it** (used to detect the same class of bug earlier next time)
---

## Step 7: Record ADR

Append to the end of `src/native-pptx/README.md`. Required fields:

```markdown
### ADR-N: Symptom title

**Problem**
Concise description of the symptom

**Root cause**
Why it happened (from the perspective of DOM processing, CSS interpretation, coordinate calculation)

**Fix**
Which file, function, and logic was changed

**Tests added**
Names of test cases added
```

---

## Completion Criteria

> **Before marking any fix as complete, always run the full compare pipeline and do a visual review.**  
> Do not commit until compare-visuals has been re-run against the current code and all slides have been inspected visually.

- [ ] `npx jest` — all tests pass
- [ ] Compare pipeline re-run: HTML generated → PPTX generated → `compare-visuals.js` executed
- [ ] `compare-report.html` has no FAILs (diff rate improved compared to before the fix)
- [ ] All slides reviewed visually; confirmed no NG diffs
- [ ] Commit targets: only changes to `.ts` / `.test.ts` / `pptx-export.md` / README
- [ ] No files from `dist/` are committed
- [ ] `slides-ci.html` is not committed
- [ ] ADR appended

---

## Post-Loop Report Format

When the improvement loop is complete, always report in the following format:

```
## Improvement Loop Report

### Changes
- Modified file(s): (e.g., src/native-pptx/dom-walker.ts)
- Summary of changes: (1-2 lines)

### Test Results
- Unit tests: N passed (N new tests added)

### Comparison Report (for local review)
Report: dist\compare-slides-ci\compare-report.html
(Open in browser to see side-by-side HTML / PPTX comparison and diff rates for all slides)

### Visual Review Results
- Slides reviewed: N
- FAILs (over threshold): N
- Visual NG (text overlap, missing elements, layout shift, etc.): N
  - Slide NNN: (description of issue)

### Commit
Branch: fix/...
Commit: (hash)
```

> **Report as a file system path, not an HTML URL.**
> Use the backslash format `dist\compare-slides-ci\compare-report.html` so the user can open it directly in Explorer or with `start`.
---

## What Never to Do

- Attempt to install LibreOffice via winget / msiexec / admin privileges
- Assume `npm run build` alone has updated the bundle
- `git add` files under `dist/`
- Skip adding a reproduction slide to `pptx-export.md` and go straight to fixing
- Leave a worktree behind (after use, run `git worktree remove` and `git worktree prune`)
- **Declare OK because the diff rate is low** (especially line-break shifts and wrapping issues do not appear in diff rate)
- **Start fixing without reading the ADR** (ignoring past decisions causes regressions)
- **Write developer local paths, business data, or confidential data directly into `pptx-export.md`** (public repository — always generalize)
- **Add or modify files unrelated to the fix** (only `.ts` / `.test.ts` / `pptx-export.md` / README changes are commit targets)
- **Create new comparison tools or helper scripts** (existing `compare-visuals.js` / `gen-pptx.js` / `diagnose-pptx.js` are sufficient; do not create new ones unless explicitly asked)