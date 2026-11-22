---
name: 'Marp Editable PPTX Development Conventions'
description: 'Development conventions for the native-pptx module of marp-to-editable-pptx. Apply when modifying dom-walker.ts or slide-builder.ts, adding tests, managing fixtures, recording ADRs, following commit conventions, or preventing regressions.'
applyTo: 'src/native-pptx/**/*.ts, src/native-pptx/test-fixtures/**, src/native-pptx/README.md'
---

# Marp Editable PPTX Development Conventions

## 🚨 CRITICAL: Known Failure Patterns (Read First)

### 1. Never Claim Fixed Without Visual Proof

A test passing does NOT mean a bug is fixed. **The only valid proof is visual comparison** — open `compare-report.html` and confirm the targeted slide **with your own eyes** (or present the before/after screenshots). The following have all been observed in this codebase:

- Unit test passes but the PPTX still renders wrong (test was measuring the wrong thing)
- Gate test passes with tolerance but the visual output is clearly broken to a human
- Agent declares "fixed" because code was changed and tests pass, but the actual render is identical to before

**Rule**: After any fix targeting a visual bug, always show the PPTX screenshot of the affected slide and state explicitly: "I compared the PPTX render and [symptom] is now [resolved/still present]."

### 2. Never Import Business Data Into Test Fixtures

When a developer shares their slide deck to demonstrate a bug:
- **DO NOT** copy, sanitize, paraphrase, or "generalize" any text content
- Domain terms, project names, numeric values, status labels, field names — even after "anonymization" — leak business context into a public repository
- The CSS/HTML **structure** is the only thing to reproduce; use only the approved vocabulary table (see Fixture Management section below)

### 3. Code Block Rendering — Partially Fixed; fontSize/Indentation OPEN

**Fixed in recent ADRs (no longer open):**
| Fix | ADR | Details |
|-----|-----|---------|
| Syntax highlighting color loss | ADR-40 | `extractCodeRuns()` rewritten to walk `<span>` tokens |
| `<pre>` inside `<li>` had no background fill | ADR-43 | Extracted as separate CodeElement shape |
| `<ul>/<ol>` inside `<blockquote>` lost bullet structure | ADR-44 | Extracted as separate ListElement |
| Code block shape missing when multiple `<pre>` per `<li>` | ADR-46 | Fixed extraction loop in dom-walker.ts |
| Code block bg taken from wrong element | ADR-46 | `<code>` bg used as fallback when `<pre>` bg is transparent |
| Code block fontSize read from wrong element (paragraph instead of `<pre>`) | ADR-46 | `extractTextStyle` now resolves `<pre>`/`<code>` specific fontSize — but the value read is still CSS-declared (pre-transform); see Still open below |

**Still open (architecturally limited):**
| Bug | Symptom | Root Cause (known) |
|-----|---------|-------------------|
| **fontSize mismatch** | PPTX code block font is ~1.5× larger than HTML | `getComputedStyle(<pre>).fontSize` returns the CSS-declared value (e.g. 24.65px) but Marp's bespoke.js auto-scaling visually shrinks it via CSS transform; ADR-46 fixed which *element* is read, but not the CSS transform gap |
| **Indentation loss** | Leading whitespace in PPTX is reduced (4sp → 2sp visually) | Combined effect of (a) wrong fontSize making each char wider, and (b) potential text inset / margin interaction in PptxGenJS when `wrap:false` |
| **Overflow** | Long code blocks overflow the shape boundary | `codeFontScale` cap in slide-builder.ts (ADR-46 Gate 2: fontSize capped so all lines fit vertically) mitigates but doesn't fix — the font is still wrong, it's just capped to fit |

**Any "fix" for code blocks MUST demonstrate improvement by showing the slide-90 PPTX screenshot.** If the font size in the PPTX still doesn't match the HTML (compare visually!), the fix is incomplete.

### 4. Auto-Scaling Detection Is Architecturally Limited

Marp's bespoke.js applies auto-scaling via CSS `transform: scale(...)` on a Shadow DOM wrapper. `getBoundingClientRect()` on the `<pre>` element returns the **post-transform** visual size, but `getComputedStyle().fontSize` returns the **pre-transform** declared value. This means:

- `computeAutoScaleFactor()` in dom-walker compares lineHeight×lines to rect height — but in the test fixture (no bespoke.js), rect == natural size, so scaleFactor is always 1.0
- In production (bespoke.js active), rect IS scaled down, but fontSize from getComputedStyle is NOT — so the mismatch exists differently

**The correct fix must either:**
1. Detect the CSS transform scale factor on the bespoke wrapper and apply it to fontSize, OR
2. Use a completely different approach (e.g., measure actual rendered text height via a probe element)

## Language Policy

All source code, comments, test case names, documentation, and ADR entries under `src/native-pptx/` must be written in **English**. No exceptions.

Exception: test fixture content that intentionally tests Japanese character rendering (e.g., mixed Japanese/English text in `pptx-export.md`) may contain Japanese, because that text is the subject under test.
## Design Principles (Violations Are Prohibited)

**Browser is the source of truth**

- Map `getComputedStyle()` and `getBoundingClientRect()` values 1:1 to PPTX
- Do not parse Marp themes, CSS selectors, or Markdown syntax
- Element-specific hardcoding is only allowed **when the browser has already rendered the result but PPTX has a structural limitation that prevents reproduction** (e.g., SVG `<foreignObject>`, slide page numbers)
- In that case, the only permitted fix is "capture the browser rendering result as a raster image"

**The render plan is the conversion source of truth**

- Clipping, visibility, grouping, and container ownership must be resolved once in the render plan.
- The PptxGenJS executor must consume only planned visible elements and associations. It must not recompute a different set from source elements.
- Plan execution and diagnostic serialization must not mutate `SlideData` or the diagnosed plan.

**Visual regression gates are independent constraints**

- Whole-slide absolute limits and page-specific baseline limits must both pass.
- A localized-region regression is a separate failure even when whole-slide percentage remains low.
- Text contracts must preserve multiplicity and reading order, including single-character and symbol content.
- Visual comparisons normalize the full rendered image to the browser/Marp reference
  canvas before pixel analysis. Never crop unequal renderer dimensions to a common
  minimum rectangle, because DPI differences then compare only part of the slide.

**Font and line geometry share one coordinate system**

- CSS transforms apply to font size, line height, padding, border, and measured geometry together.
- Do not subtract untransformed padding or border from a transformed browser rectangle.
- When PowerPoint cannot preserve browser line layout inside one text box, retain measured per-line editable text boxes rather than relying on PowerPoint wrapping or baseline inference.
- Apply the renderer font compensation consistently to every editable text path,
  including table cells. A table fallback branch must not bypass the profile.
- Reserve a small single-line descent allowance so PowerPoint glyph bounds do
  not touch a browser-sized box edge or adjacent rule.
- Narrow glyph-only text boxes use centered/forward width slack before font
  shrinking. Arrows and single symbols must remain editable and fully visible.
- Tables own their measured row heights and column totals. Renderer compensation
  must change proportions, not expand the table beyond its browser width.
- Ordered-list text starts at the browser-rendered indent, not PptxGenJS's
  default 27pt margin. List-specific CJK compensation and width slack apply
  together to prevent one-character wrap. A list's geometry retains the full
  horizontal flow width from its HTML parent; an intrinsic text-only rect must
  not be converted into a large right inset when embedded in a card.
- Code-block CSS padding is part of editable text geometry and must become PPTX
  inset rather than being discarded.
- Container-embedded text honors CSS padding as content-box geometry even when
  the text element and its visual container share the same browser bounds.

**CSS paint follows browser ownership and side geometry**

- Preserve CSS border geometry per side. A visible top border is not permission
  to synthesize a four-sided frame, and a 6px left accent must remain exactly
  on the left instead of becoming part of a uniform outline.
- `::before` and `::after` are painted descendants, even when they are not DOM
  nodes. Extract absolute/fixed-position visual pseudo-elements across the
  section tree, but keep generated pseudo text from being mistaken for
  decorative graphics.
- Positioned elements with explicit CSS `z-index` retain that stacking level
  across nested container boundaries. Positive overlays paint after ordinary
  content and remain ordered by their stacking index.
- An SVG rendered as an image owns all of its visual text. SVG descendant text
  must not also be emitted as editable paragraphs.
- Give pseudo-elements synthetic DOM paths. They must never become the source
  owner of ordinary descendant text just because they were generated by the
  same host element.
- An image inside an `overflow:hidden` or `overflow:clip` ancestor is a clipped
  viewport, not an oversized standalone image. Rasterize the visible viewport
  when native embedding cannot express the crop exactly.

**Cancellation and performance are part of the conversion contract**

- Long-running browser, screenshot, and file-resolution operations must observe cancellation without converting the abort into a rasterization error.
- Shared local images must be resolved once per export.
- Conversion tools must expose stage timings so end-to-end performance can be measured instead of inferred from planner-only benchmarks.
- PR validation must generate the full 96-slide fixture and run extraction/XML contracts without conditional skips.

**DOM order remains a content contract**

- Badge, text, list, emoji, and container output must follow DOM reading order, not extraction phase order.
- A list containing inline badges may be split into positioned list chunks and badge shapes so source order remains editable.
- Browser-context functions must not close over Node-only imports; serialized functions must use browser globals explicitly.

**Cancellation and DOM transformation boundaries are end-to-end contracts**

- Cancellation after any uncancellable operation starts must prevent publication of its result. Write output through a temporary sibling and replace the destination only after the operation and all subsequent abort checks succeed.
- If an uncancellable publication operation can complete after cancellation, roll
  it back after it resolves. Preserve any pre-existing destination from a sibling
  backup and treat rollback failure as an error rather than reporting cancellation
  as a successful no-output state.
- Transform ancestry continues from a ShadowRoot through its host.
- Badge-adjacent list markers retain PowerPoint list semantics and CSS list style; measured text placement must not replace list structure with literal marker text.
- OOXML text contracts validate against extracted browser data, never against the render plan built from that same data.

**The 96-slide fixture must be an honest acceptance surface**

- Every visible source element must fit within the Marp canvas unless the fixture explicitly exists to test clipping; fixture source overflow is not an acceptable baseline.
- Local fixture assets must exist and remote assets are prohibited.
- Stable fixture IDs come from `fixture-manifest.json`; page numbers and display titles are not interchangeable identifiers.
- The production-like bespoke active-slide path must preserve the same ordered text contract as static extraction.
- OOXML semantics must be asserted from extracted element types: tables remain tables, lists retain markers, code remains editable, and hyperlinks remain relationships.
- Export extracts the final Marp state: inactive fragments are made visible before structure extraction, matching the comparison renderer's complete-slide view.
- CSS gradients and rounded strong-highlight backgrounds use editable underlay shapes at the measured inline bounds; editable text must not carry a PowerPoint full-glyph highlight approximation for those cases.

**Generated native bundle inputs have an ordering contract**

- Generate `dom-walker-script.generated.ts` before building
  `lib/native-pptx.cjs` whenever local tasks rebuild both artifacts.
- A build that bundles first can consume a stale DOM walker; CI, VS Code tasks,
  and developer instructions must use the same dependency order.

**Structural fidelity over visual approximation**

When PPTX has a rendering limitation (e.g., solid-color-only text backgrounds, no CSS transparency):

- HTML-specified structural elements — code highlights, borders, bullets — are **always** rendered in PPTX, even when the result is visually imperfect.
- Do **not** suppress structural elements based on "the composited result looks wrong in this specific context". That is an ad-hoc visual approximation, not a structural decision.
- Known imperfections must be documented in the **"Known limitations" section of `src/native-pptx/README.md`**, not worked around with per-element special cases.
- Exception: suppress only when the element would be **entirely invisible** — i.e., its composited color is indistinguishable from the slide's visual background — AND a general rule (not a per-case heuristic) can describe it. Example: near-white (`r,g,b > 200`) highlight on a full-slide dark bg image (≥80% slide width, CSS bg is white) would be invisible against the image → suppress by that general rule (see `visualBgMayBeDark` in `slide-builder.ts`).

**Layout wrappers and geometry**

- A zero-sized bounding box does not by itself mean an element is invisible.
- Elements with computed `display: contents` have no box but participate
  through their rendered child nodes. Treat them as transparent wrappers and
  traverse `childNodes` in DOM order before applying zero-size filtering.
- Direct text nodes inside `display: contents` participate in layout as
  anonymous boxes. Measure each non-empty text node with a DOM `Range` and emit
  a paragraph before continuing to later element nodes.
- Inline element children skipped by the normal block walker must also be
  emitted from their rendered bounding boxes so transparent wrappers do not
  lose inline content in a non-grid/non-flex parent.

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
- Preserve the 96-slide fixture unless an explicit product decision removes a
  documented scenario. Add combinations instead of replacing existing slides.
- Visual pixel rate is not a content oracle. DOM-walker changes must preserve
  the full-fixture browser Text-node → `SlideData` contract.
- Cover `display` and node combinations with a matrix, especially parent layout,
  wrapper depth, inline/block/text children, visibility overrides, and
  zero-sized boxes.
- Update page-specific visual baselines intentionally; do not raise the global
  threshold to absorb a new structural difference.
- Keep the visual gates aligned: compare/quick-diff use 3% WARN and 6% FAIL;
  the PowerPoint all-slide gate additionally allows 1 percentage point above
  each page baseline.
- Do not use fixed sleeps as render readiness. Wait for document, fonts,
  images, and stable slide geometry.
- Required rasterization failures must fail the export; never convert an
  incomplete screenshot path into success.
- Review both whole-slide diff and localized region diff. A low whole-slide
  percentage must not clear a concentrated badge, icon, or table regression.
- PowerPoint-specific compensation belongs in `renderer-profile.ts`, not in
  element placement branches.
- `SlideData` must pass through a testable PPTX render plan before PptxGenJS
  side effects. Plan nodes must preserve source identity unless a structural
  change is intentional.
- Preserve the complete contract chain: browser rendered content →
  `SlideData` → PowerPoint plan → PPTX XML → visual render.

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
