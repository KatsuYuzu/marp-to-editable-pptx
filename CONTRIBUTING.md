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

2. **Add a minimal fixture slide** to `src/native-pptx/test-fixtures/pptx-export.md` that reproduces the problem. Then add the corresponding `<tr>` row(s) for the new slide(s) in `README.md` — the `compare-NNN.png` row inside `<details>`.

   > **⚠️ STOP — Fixture Content Safety Gate**  
   > If you (or an AI assistant) can see the content of your actual slide right now, **do not copy, adapt, or "sanitize" it**.  
   > Business data cannot be made safe by substitution — domain meaning survives sanitization.  
   > **Write all fixture text from scratch using only these building blocks**: `Label-A`, `Cat-B`, `Item-N`, `Tag-C`, `Group A`, `val-N`, `/uu`, `Alpha beta gamma delta`, `Zeta nu eta theta kappa`.  
   > See `.github/instructions/marp-editable-pptx.instructions.md` — Fixture Management for the full approved vocabulary.

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
docs(readme): add compare image row after adding fixture
```

Scope = the filename without extension (`dom-walker`, `slide-builder`, `index`, `utils`, `compare-visuals`, etc.).

## Pull requests

- Branch naming: `fix/description-in-kebab-case` or `feat/description-in-kebab-case`
- One PR per fix or feature
- Do **not** commit `dist/` files or `slides-ci.html`

## Releasing

Releases are fully automated once a `v*` tag is pushed. The manual steps are:

1. On the release/fix branch, bump the version with
   `npm version patch --no-git-tag-version` (use `minor` or `major` when appropriate,
   following [semver](https://semver.org/)). This updates both `package.json` and
   `package-lock.json`.
2. Update `CHANGELOG.md` with the new version, date, user-visible changes, and any
   real-machine verification limitations. The changelog is shipped with the extension
   and serves as its release notes.
3. Run type checking, tests, and a production build. Package a VSIX with
   `npx vsce package` and check its version before publishing.
4. Commit and push the release/fix branch, then merge it into `main` through a PR.
   Do not include machine-specific debug settings. Use `Refs #N` rather than
   `Fixes #N` when reporter confirmation is still pending.
5. Fetch the merged commit, confirm its `package.json` version matches the intended
   tag, then tag that commit and push only the tag:
   ```sh
   git fetch origin
   git tag -a vX.Y.Z origin/main -m "Release vX.Y.Z"
   git push origin vX.Y.Z
   ```

Pushing the tag triggers `release.yml`, which runs type check and tests, then publishes to the VS Code Marketplace via `vsce publish`.
Do not push the release tag before the PR is merged. After publication, ask affected
reporters to update the installed extension and retry; cloning a branch is not required.

**Prerequisite:** The `VSCE_PAT` secret must be set in repository settings (a Personal Access Token from [dev.azure.com](https://dev.azure.com) with *Marketplace > Manage* permission).

## CI workflows and GitHub Pages

### Workflow overview

| Workflow | Trigger | What it does |
|---|---|---|
| `ci.yml` | Push or PR to any branch | Type check + full test suite |
| `screenshots.yml` | Push to `main` with changes in `src/native-pptx/**` or `scripts/gen-html-screenshots.js`; any `v*` tag push; or manual `workflow_dispatch` | Generates comparison images → publishes to GitHub Pages |
| `release.yml` | Push of any `v*` tag | Type check + tests → publishes to VS Code Marketplace |

### How screenshots and GitHub Pages work

`screenshots.yml` is the pipeline that keeps the comparison images in `README.md` up to date:

1. **HTML → PNG** via Puppeteer (Chrome headless)
2. **PPTX → PDF → PNG** via LibreOffice + pdftoppm (150 dpi)
3. **Side-by-side** via ImageMagick (`compare-NNN.png`)
4. Publishes all three sets to the **`gh-pages` branch** under `screenshots/`

`README.md` references those images with absolute GitHub Pages URLs:
```
https://katsuYuzu.github.io/marp-to-editable-pptx/screenshots/compare-NNN.png
```

**Key points for contributors and AI:**

- When you add a new fixture slide, add the `<img>` tag for `compare-NNN.png` to `README.md`'s `<details>` block — CI will generate the actual image on the next qualifying push. You do **not** need to generate or commit the PNG yourself.
- CI uses **LibreOffice** for PPTX→PNG. Local development uses **PowerPoint COM** (`compare-visuals.js`). The outputs are not pixel-identical; slight rendering differences are expected and normal.
- The RMSE threshold in CI is 0.20. Exceeding it logs a warning but does **not** fail the job. Only catastrophic rendering failures (missing slides, solid color blocks) would cause a visual regression worth investigating.
- To force a screenshot refresh without a code change (e.g., after updating only the fixture Markdown), use `workflow_dispatch` from the Actions tab.

## License

By contributing you agree your changes are released under the [MIT License](LICENSE).
