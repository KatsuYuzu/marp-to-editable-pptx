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

### 3. Code Block Bugs Are OPEN (Not Fixed)

Slides 87, 87b, 90, 92 code block issues remain **unresolved**:

| Bug | Slides | Status | What Was Done (Insufficient) |
|-----|--------|--------|------------------------------|
| fontSize ~1.5× too large | 87, 90 | **OPEN** | Gate 2 caps overflow but font still wrong |
| Indentation halved | 87, 90 | **OPEN** | `extractCodeRuns` preserves text but wrong fontSize → wrong char width |
| Overflow/clipping | 87b, 90 | **Mitigated only** | `codeFontScale` cap prevents spill, not a fix |

**Any work on code blocks MUST show slide-87/90 PPTX screenshots compared to HTML.**

### 4. Auto-Scaling Detection Is Architecturally Limited

Marp bespoke.js applies `transform: scale(...)` in Shadow DOM:
- `getComputedStyle(<pre>).fontSize` → **pre-transform** CSS value (24.65px)
- `getBoundingClientRect().height` → **post-transform** visual size

In test fixtures (no bespoke.js): no transform → no mismatch → tests pass → bug invisible.

**Correct fix must**: detect transform scale factor, use rendered measurement, or make test fixture use bespoke.js.

---

## Architecture

```
VS Code Extension (extension.ts)
  → marpCli: MD → HTML (with bespoke.js in production)
  → generateNativePptx (index.ts)
    → Puppeteer loads HTML
    → page.evaluate(extractSlides) — dom-walker.ts in browser VM
    → SlideData[] extracted
    → resolveImageUrls (rasterization for SVG/filtered images)
    → buildPptx(slides) — slide-builder.ts
    → PptxGenJS.write() → Buffer
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

### Why Tests Don't Catch It

Test fixture = `npx marp --html` → static HTML, no bespoke.js, no transform → tests pass.

### Current Mitigations

1. `computeAutoScaleFactor()`: lineHeight×lines vs rect.height → works in prod, not in tests
2. Gate 2 `codeFontScale`: caps fontSize to fit shape → prevents overflow, font still wrong
3. Gate 3 visual-regression.test.ts: 8% pixel threshold → current ~5.5% on slide 90 passes

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

```powershell
npx jest                                         # All (~346)
npx jest --testNamePattern "Gate 2"              # Font overflow gate
npx jest src/native-pptx/visual-regression.test  # Gate 3 (needs PowerPoint)
```

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

---

## Quick Reference: Fix Location

| Symptom | Fix in |
|---------|--------|
| Text not extracted / extra | `dom-walker.ts` |
| Wrong font size | dom-walker extraction → slide-builder conversion |
| Wrong color/fill | `slide-builder.ts` |
| Missing image | `index.ts` |
| Indentation lost | `extractCodeRuns` whitespace + PptxGenJS margin |
| Overflow | `slide-builder.ts` (shape sizing, fontSize cap) |
| Auto-scaling not detected | `dom-walker.ts` — architecturally limited |

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
