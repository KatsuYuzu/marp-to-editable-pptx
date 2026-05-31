---
name: marp-editable-pptx
description: 'System skill for the marp-to-editable-pptx VS Code extension. Covers the full architecture (extension → marpCli → Puppeteer → dom-walker → slide-builder → PptxGenJS), the visual fidelity improvement loop (gen-pptx + compare-visuals + PowerPoint COM), known open bugs (code block fontSize/indentation, auto-scaling detection), fixture safety rules, ADR recording, and critical failure patterns that have caused false "fixed" claims. Use when: modifying dom-walker.ts or slide-builder.ts, diagnosing PPTX rendering issues, running compare-visuals, adding test fixtures, fixing code block bugs, reviewing PRs, or onboarding to this codebase. Trigger words: "PPTX output", "slide layout", "text missing", "compare-visuals", "visual diff", "code block", "font size", "indentation", "auto-scaling", "marp", "pptx", "dom-walker", "slide-builder", "fixture", "ADR".'
argument-hint: 'Symptom description (e.g., "Slide 90 font size too large") or task (e.g., "fix code block indentation")'
---

# marp-to-editable-pptx — System Reference Skill

## 🚨 Critical Failure Patterns (Read This First)

### 1. Never Claim "Fixed" Without Visual Proof

**Observed repeatedly in this repository**: Agent modifies code, tests pass, declares "fixed" — but the PPTX still renders identically to the broken state.

Root causes of false "fixed" claims:
- Unit test passes but measures the wrong property (e.g., testing font cap, not font match)
- Gate test tolerance hides the problem (e.g., `375.000006 ≤ 376` passes, font still 1.5× too large)
- Test fixture HTML doesn't use bespoke.js → auto-scaling bugs invisible in unit tests
- `codeFontScale` cap prevents overflow but doesn't fix fontSize mismatch — mitigation ≠ resolution
- Agent runs `npx jest` (344 tests pass) and reports success without looking at PPTX output

**Mandatory rule**: After any visual bug fix, show PPTX screenshot and state: "Slide N PPTX render — [symptom] is [resolved/still present/unchanged]."

### 2. Never Import Business Data Into Fixtures

When a developer shares slides to demonstrate a bug:
- DO NOT copy, sanitize, paraphrase, or generalize ANY text content
- Use ONLY approved vocabulary (Label-A, Cat-B, Alpha beta gamma, val-N, etc.)
- The CSS/HTML **structure** is the only thing to reproduce; text meaning is irrelevant

### 3. Code Block Issues — Partial Status

**Fixed in recent ADRs (no longer open):**
| Fix | ADR | Details |
|-----|-----|---------|
| Syntax highlighting color loss | ADR-40 | `extractCodeRuns()` rewritten to walk `<span>` tokens |
| `<pre>` inside `<li>` had no background fill | ADR-43 | Extracted as separate CodeElement shape |
| `<ul>/<ol>` inside `<blockquote>` lost bullet structure | ADR-44 | Extracted as separate ListElement |
| Code block shape missing when multiple `<pre>` per `<li>` | ADR-46 | Fixed extraction loop in dom-walker.ts |
| Code block bg taken from wrong element | ADR-46 | `<code>` bg used as fallback when `<pre>` bg is transparent |
| Code block fontSize inherited from surrounding paragraph | ADR-46 | `extractTextStyle` now resolves `<pre>`/`<code>` specific fontSize |

**Still architecturally limited (OPEN):**
| Bug | Slides | Status | Root Cause |
|-----|--------|--------|------------|
| fontSize ~1.5× too large vs HTML | 90, 91 | **OPEN** (ADR-46) | `getComputedStyle(<pre>).fontSize` is pre-transform; bespoke.js `transform:scale()` not reflected in test fixture |
| Indentation appears reduced | 90 | **OPEN** | Depends on fontSize correctness — char width wrong at inflated size |
| Overflow/clipping | 87b | **Mitigated** | `codeFontScale` cap prevents spill; `computeAutoScaleFactor`+`applyAutoScale` applied but ineffective in fixture (no bespoke.js → scale=1 always) |

**Any work on code block fontSize MUST demonstrate improvement on slide-90/91 PPTX screenshots.**

### 4. Auto-Scaling Detection Is Architecturally Limited

Marp bespoke.js applies `transform: scale(...)` in Shadow DOM:
- `getComputedStyle(<pre>).fontSize` → **pre-transform** CSS value (24.65px)
- `getBoundingClientRect().height` → **post-transform** visual size

In test fixtures (no bespoke.js): no transform → no mismatch → tests pass → bug invisible.

`computeAutoScaleFactor()` + `applyAutoScale()` (dom-walker.ts) attempt to detect this at extraction time, but return scaleFactor=1 in fixtures. See ADR-46.

**Correct fix must**: detect the CSS transform scale factor on the bespoke wrapper and apply it to fontSize.

---

## Architecture

```
VS Code Extension (extension.ts)
  → marpCli: MD → HTML (with bespoke.js in production)
  → generateNativePptx (index.ts)
    → Puppeteer loads HTML
    → inject CSS freeze (animation:none, transition:none) — ADR-41
    → page.evaluate(extractSlides) — dom-walker.ts in browser VM
    → SlideData[] extracted
    → resolveImageUrls (rasterization for SVG/filtered images)
    → buildPptx(slides)
      → groupAdjacentTextElements() — ADR-39
      → slide-builder.ts → PptxGenJS.write() → Buffer
```

### File Responsibilities

| File | Role | Modify when |
|------|------|-------------|
| `dom-walker.ts` | Extract SlideData from browser DOM | Text missing/extra, wrong coordinates, wrong fontSize |
| `slide-builder.ts` | SlideData → PptxGenJS API calls | PPTX format, colors, fonts, overflow |
| `index.ts` | Pipeline, Puppeteer lifecycle | Image rasterization, browser setup |
| `utils.ts` | Unit conversion (px→inch, rgb→hex) | Formula bugs |
| `types.ts` | Shared TypeScript interfaces | New element properties |
| `extension.ts` | VS Code command, marpCli | UX flow |

### Key Functions Added in Recent ADRs

| Function | File | Added in | Purpose |
|----------|------|----------|---------|
| `extractCodeRuns()` | dom-walker.ts | ADR-40 | Extract syntax-highlighted `<span>` runs from `<pre><code>`; preserves leading whitespace (ADR-45) |
| `computeAutoScaleFactor()` | dom-walker.ts | ADR-40/46 | Detect over-height code blocks via lineHeight×lines vs rectHeight; limited by bespoke.js absence in fixtures |
| `applyAutoScale()` | dom-walker.ts | ADR-40/46 | Scale down run.fontSize × scaleFactor when < 1; mutates runs in-place |
| `groupAdjacentTextElements()` | slide-builder.ts | ADR-39 | Merge adjacent text shapes into one text box |
| `cssListStyleToNumberStyle()` | slide-builder.ts | ADR-42 | Map CSS list-style-type → PptxGenJS number style |
| `cssListStyleToBulletChar()` | slide-builder.ts | ADR-42 | Map CSS list-style-type → Unicode bullet char |

### Build Chain (⚠️ npm run build does NOT rebuild the native-pptx bundle)

```powershell
node src/native-pptx/scripts/generate-dom-walker-script.js   # after dom-walker.ts
node src/native-pptx/scripts/build-native-pptx-bundle.js     # after any .ts in native-pptx
# → lib/native-pptx.cjs (used by gen-pptx.js and visual-regression.test.ts)
```

### Conversion Formulas

```
pxToInches(px) = px / 96
pxToPoints(px) = px * 0.75
PptxGenJS: x/y/w/h in inches, fontSize in points (72pt = 1 inch)
```

---

## The Auto-Scaling Problem (Code Block Root Cause)

### Marp Auto-Scaling (Production)

1. marpCli renders HTML with bespoke.js
2. Bespoke.js detects content overflow > 720px
3. Wraps in Shadow DOM with `transform: scale(factor)`
4. CSS `fontSize` unchanged — only visual transform applies

### What dom-walker Extracts

| Property | Source | Transform-aware? |
|----------|--------|-----------------|
| `el.style.fontSize` | `getComputedStyle(<pre>).fontSize` | **NO** |
| `el.width/height` | `getBoundingClientRect()` | **YES** |
| `run.fontSize` | `getComputedStyle(parent).fontSize` | **NO** |

**Result**: Shape box correct (transformed size) but font too large (untransformed value).

### Current Handling

1. `computeAutoScaleFactor()`: lineHeight×lines vs contentHeight → detects overflow, but returns 1 in fixture (no transform)
2. `applyAutoScale()`: scales run.fontSize × scaleFactor when < 1; mutates runs in-place
3. Gate 2 `codeFontScale`: caps fontSize to fit shape → prevents overflow, font still wrong if computeAutoScaleFactor returns 1
4. Gate 3 `visual-regression.test.ts`: 8% pixel threshold for slides 87b + 90 (CODE_BLOCK_SLIDES); runs with PowerPoint COM

### Why Tests Don't Catch It

Test fixture = `npx marp --html` → static HTML, no bespoke.js, no transform → `computeAutoScaleFactor` returns 1 → `applyAutoScale` is a no-op → tests pass.

---

## Visual Diff Loop (Full Pipeline)

```powershell
# 1. Rebuild
node src/native-pptx/scripts/generate-dom-walker-script.js
node src/native-pptx/scripts/build-native-pptx-bundle.js

# 2. HTML (if pptx-export.md changed)
npx marp src/native-pptx/test-fixtures/pptx-export.md `
  --html --allow-local-files `
  --output src/native-pptx/test-fixtures/slides-ci.html

# 3. Unit tests
npx jest

# 4. PPTX
node src/native-pptx/tools/gen-pptx.js `
  src/native-pptx/test-fixtures/slides-ci.html dist/slides-ci.pptx

# 5. Compare (needs PowerPoint)
node src/native-pptx/tools/compare-visuals.js `
  src/native-pptx/test-fixtures/slides-ci.html dist/slides-ci.pptx

# 6. Open dist/compare-slides-ci/compare-report.html
```

### Diff Rate Is NOT the Acceptance Criterion

Typography differences trigger FAIL thresholds — acceptable. Line-break shifts can have 0% diff. **Always classify visually.**

---

## Tests

There are no "Gate 1" unit tests by that name. ADR-46 introduced a 3-level testing strategy:

| Level | What | Command |
|-------|------|--------|
| Extraction tests (ADR-46) | dom-walker correctly extracts code elements — shape count, bg, fontSize isolation | `npx jest` (dom-walker.test.ts) |
| Gate 2 (ADR-46) | slide-builder's `codeFontScale` correctly caps fontSize to prevent overflow | `npx jest --testNamePattern "Gate 2"` |
| Gate 3 | PowerPoint visual diff — structural overflow produces >15% pixel diff | `npx jest src/native-pptx/visual-regression.test` |

**Extraction tests (dom-walker.test.ts, `ADR-46:` describe blocks)**:
- `ADR-46: multiple code blocks in list items all produce separate code elements` — each `<pre>` in `<li>` emits its own CodeElement
- `ADR-46: code block background from <code> element when <pre> is transparent` — fallback bg detection
- `ADR-46: code block fontSize must not inherit parent paragraph size` — font isolation
- `ADR-46: code block fontSize must be adjusted when auto-scaling shrinks the box` — `computeAutoScaleFactor` + `applyAutoScale` behavior

**Gate 2 (`slide-builder.test.ts`, describe: `ADR-46 Gate 2`)**: Tests that `codeFontScale` caps fontSize when `lineCount × fontSize × 1.2 > shapeHeight`. Validates physical overflow prevention — NOT fontSize fidelity.

**Gate 3 (`visual-regression.test.ts`)**: Generates PPTX from `slides-ci.html`, exports slides to PNG via PowerPoint COM, pixel-diffs HTML screenshots vs PPTX. Covers `CODE_BLOCK_SLIDES` = slides 87b + 90. Threshold: 8% pixel diff. Slide 91 is NOT currently in Gate 3.

---

## Fixture Rules

- Approved vocabulary ONLY: `Label-A`, `Cat-B`, `val-N`, `Alpha beta gamma`
- Never: domain terms, real numbers, status labels, workflow steps
- Scope `<style>` to `section`
- Run compare-visuals for ALL slides after adding
- Add README `<tr>` row in same commit

---

## ADR Log

`src/native-pptx/README.md` → "Bug fix and decision log". Required on every `.ts` fix.

Latest ADRs:
- ADR-39: Text grouping (adjacent text → single box)
- ADR-40: Code block syntax highlighting (colors)
- ADR-41: CSS animation/transition freeze on Puppeteer capture
- ADR-42: `list-style-type` mapping
- ADR-43: `<pre>` inside `<li>` → separate CodeElement shape
- ADR-44: `<ul>/<ol>` inside `<blockquote>` → separate ListElement
- ADR-45: Regression test fixtures for ADR-43/44 (slides 87-89 in pptx-export.md); fixture-only, no `.ts` change
- ADR-46: Code block shape loss (multi-`<pre>`-per-`<li>`), bg fallback, fontSize inheritance fix, `codeFontScale` Gate 2 test, Gate 3 visual regression test; font size fidelity vs bespoke.js transform remains OPEN

---

## Quick Reference: Fix Location

| Symptom | Fix in |
|---------|--------|
| Text not extracted / extra | `dom-walker.ts` |
| Wrong font size | dom-walker extraction → slide-builder conversion |
| Syntax highlighting colors wrong | `extractCodeRuns()` in dom-walker.ts (fixed ADR-40) |
| Wrong color/fill | `slide-builder.ts` |
| Missing image | `index.ts` |
| Indentation lost | `extractCodeRuns` whitespace + PptxGenJS margin |
| Overflow | `slide-builder.ts` (shape sizing, fontSize cap) |
| Auto-scaling not detected | `dom-walker.ts` — architecturally limited (ADR-46 OPEN) |
| Adjacent text elements not merged | `groupAdjacentTextElements()` in slide-builder.ts (fixed ADR-39) |
| list-style-type wrong | `cssListStyleToNumberStyle()` / `cssListStyleToBulletChar()` (fixed ADR-42) |
| `<pre>` in list loses bg fill | `dom-walker.ts` — ADR-43 pattern (fixed) |
| paginate:hold slides dropped | `dom-walker.ts` — SVG grouping (ADR-33) |

---

## What Never to Do

- Declare "fixed" without PPTX screenshot proof
- Import business data into fixtures
- Assume `npm run build` rebuilds native-pptx bundle
- Judge by diff rate alone
- Skip ADR log before fixing
- Install LibreOffice locally
- Create new tools without being asked
- Commit `dist/` or `slides-ci.html`
