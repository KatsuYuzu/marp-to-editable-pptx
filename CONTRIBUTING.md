# Contributing

Thank you for your interest in contributing to **Marp to Editable PPTX**!

## Quick orientation

| File | Role |
|---|---|
| `src/native-pptx/dom-walker.ts` | Extracts layout from browser DOM |
| `src/native-pptx/slide-builder.ts` | Maps extracted data to PptxGenJS API calls |
| `src/native-pptx/index.ts` | Pipeline controller (browser lifecycle, image rasterization) |
| `src/native-pptx/README.md` | Architecture details and ADR log |

## Setup

```sh
npm install
npm run build
npm test
```

> **Note:** `npm run build` only builds the VS Code extension bundle. If you change `dom-walker.ts`, you must also run the extra scripts listed in step 5 below.

## Reporting bugs

Use the [bug report template](.github/ISSUE_TEMPLATE/bug_report.md).  
A minimal reproduction `.md` file is the most useful thing you can provide — please remove any confidential content before sharing.

> **Font/wrap differences** between the browser and PowerPoint are a [known limitation](README.md#why-this-extension) rather than a bug in most cases. Still worth reporting if the difference is severe.

## Making changes

1. **Read the ADR log first** — `src/native-pptx/README.md` documents past decisions and fixed bugs. Skipping this step is the most common cause of regressions.

2. **Add a minimal fixture slide** to `src/native-pptx/test-fixtures/pptx-export.md` that reproduces the problem. Then update the slide count in **two places**:
   - `README.md` — the `compare-NNN.png` line in `<details>` and the `All slide comparisons (N slides)` count
   - `src/native-pptx/README.md` — the "Canonical test deck" section

3. **Verify the fixture reproduces the issue** before touching any code:
   ```sh
   npx marp src/native-pptx/test-fixtures/pptx-export.md --html --allow-local-files \
     --output src/native-pptx/test-fixtures/slides-ci.html
   node src/native-pptx/tools/gen-pptx.js \
     src/native-pptx/test-fixtures/slides-ci.html dist/compare-out.pptx
   node src/native-pptx/tools/compare-visuals.js \
     src/native-pptx/test-fixtures/slides-ci.html dist/compare-out.pptx
   # open dist/compare-slides-ci/compare-report.html
   ```
   Check **all slides visually** — a low diff percentage does not mean no problem (line-wrap shifts are invisible to pixel diff).

4. **Fix the bug**, then add a regression test in `dom-walker.test.ts` or `slide-builder.test.ts`. Test names must be in English.

5. **Rebuild** if you changed `dom-walker.ts`:
   ```sh
   node src/native-pptx/scripts/generate-dom-walker-script.js
   node src/native-pptx/scripts/build-native-pptx-bundle.js
   ```

6. **Re-run the visual comparison** (same commands as step 3) to confirm the fix improved the diff and introduced no regressions.

7. **Record an ADR** — append to `src/native-pptx/README.md` under "Bug fix and decision log" with five fields, all written in **English**: problem, root cause, fix, test name, and why it wasn't caught earlier.

## AI-assisted development (GitHub Copilot)

> **Note:** This project is developed exclusively with AI assistance (GitHub Copilot). All source code, tests, and documentation are produced through human-AI collaboration — there is no hand-written code. When contributing, you are encouraged to use AI tools too.

This repository ships with Copilot customizations that encode the workflow above:

- **Skill** — `.github/skills/marp-pptx-visual-diff/SKILL.md`
- **Instructions** — `.github/instructions/marp-editable-pptx.instructions.md`

These files load project conventions (browser-as-truth principle, ADR requirement, language policy, degression-prevention checklist) directly into the AI's context. They are the main reason AI suggestions stay consistent across sessions.

## Commit style

Follow [Conventional Commits](https://www.conventionalcommits.org/):

```
fix(dom-walker): prevent pre-render text nodes from appearing twice
feat(slide-builder): support CSS custom properties in fill color
docs(readme): update slide count after adding fixture
```

Scope = the filename without extension (`dom-walker`, `slide-builder`, `index`, `utils`, `compare-visuals`, etc.).

## Pull requests

- Branch naming: `fix/description-in-kebab-case` or `feat/description-in-kebab-case`
- One PR per fix or feature
- Do **not** commit `dist/` files or `slides-ci.html`

## License

By contributing you agree your changes are released under the [MIT License](LICENSE).
