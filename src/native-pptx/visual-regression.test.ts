/**
 * Gate 3: Visual regression test — pixel-level fidelity comparison.
 *
 * Generates PPTX from the test fixture HTML, renders both HTML and PPTX
 * to PNG screenshots, then compares them via pixelmatch.
 *
 * Requirements:
 * - Chrome/Chromium or Microsoft Edge installed (auto-detected)
 * - PowerPoint desktop (COM automation) on Windows
 * - lib/native-pptx.cjs built (npm run build)
 *
 * Skips automatically if PowerPoint, Chromium-based browser, or the generated
 * fixture HTML is unavailable.
 */
import { spawnSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { detectBrowserPath } from './browser'
import visualBaselines from './test-fixtures/visual-baselines.json'
import { evaluateVisualGate } from './visual-gate'

// --- Environment detection (run once at module load) ---

function hasPowerPoint(): boolean {
  if (process.platform !== 'win32') return false
  try {
    const r = spawnSync(
      'powershell',
      [
        '-NoProfile',
        '-Command',
        'New-Object -ComObject PowerPoint.Application | Out-Null; Write-Output "OK"',
      ],
      { encoding: 'utf-8', timeout: 15000 },
    )
    return r.stdout?.includes('OK') ?? false
  } catch {
    return false
  }
}

const CHROME_PATH = detectBrowserPath('auto', process.env.CHROME_PATH)
const HAS_POWERPOINT = hasPowerPoint()
const PROJECT_ROOT = path.resolve(__dirname, '../..')
const BUNDLE_PATH = path.join(PROJECT_ROOT, 'lib', 'native-pptx.cjs')
const HAS_BUNDLE = fs.existsSync(BUNDLE_PATH)
const FIXTURE_HTML_PATH = path.join(
  PROJECT_ROOT,
  'src',
  'native-pptx',
  'test-fixtures',
  'slides-ci.html',
)

const CAN_RUN = !!(
  CHROME_PATH &&
  HAS_POWERPOINT &&
  HAS_BUNDLE &&
  fs.existsSync(FIXTURE_HTML_PATH)
)

let WIDTH = visualBaselines.width
let HEIGHT = visualBaselines.height

// --- Pixel diff thresholds ---
// These must be STRICT for code block overflow detection.
// Normal font rendering diff (Chrome vs PowerPoint) is ~2-5%.
// Code block overflow causes >15% diff because text spills outside the shape.
const FAIL_THRESHOLD = visualBaselines.absoluteFailPct / 100
const BASELINE_TOLERANCE_PCT = visualBaselines.baselineTolerancePct
const MAX_REGION_TOLERANCE_PCT = visualBaselines.maxRegionTolerancePct

// Baseline captured with the current PowerPoint/Edge renderer for all fixture
// slides. The tolerance covers documented COM/font-renderer variance while
// still detecting regressions that stay below the absolute 6% cutoff.
const VISUAL_BASELINES_PCT = visualBaselines.slides
const MAX_REGION_BASELINES_PCT = visualBaselines.maxRegionPct

// --- Helper: generate PPTX from HTML ---
async function generatePptx(
  htmlPath: string,
  outputPath: string,
): Promise<void> {
  // Use gen-pptx.js as a child process to avoid Jest VM restrictions
  // (pptxgenjs uses dynamic import() which requires --experimental-vm-modules)
  const genScript = path.join(
    PROJECT_ROOT,
    'src',
    'native-pptx',
    'tools',
    'gen-pptx.js',
  )
  const result = spawnSync(
    'node',
    [genScript, htmlPath, outputPath, CHROME_PATH!],
    {
      encoding: 'utf-8',
      timeout: 120000,
      cwd: PROJECT_ROOT,
    },
  )
  if (result.status !== 0) {
    throw new Error(
      `gen-pptx.js failed (exit ${result.status}): ${result.stderr || result.stdout}`,
    )
  }
}

// --- Helper: screenshot HTML slides via Puppeteer ---
async function screenshotHtmlSlides(
  htmlPath: string,
  outDir: string,
  slideIndices: number[],
): Promise<Map<number, string>> {
  const puppeteer = require('puppeteer-core')
  const browser = await puppeteer.launch({
    executablePath: CHROME_PATH,
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu'],
  })

  const results = new Map<number, string>()
  try {
    const page = await browser.newPage()
    await page.setViewport({ width: WIDTH, height: HEIGHT })
    await page.goto(pathToFileURL(htmlPath).href, {
      waitUntil: 'networkidle0',
    })
    const canvas = await page.evaluate(() => {
      const svg = document.querySelector('svg[data-marpit-svg]')
      const viewBox = svg?.getAttribute('viewBox')?.split(/\s+/).map(Number)
      return viewBox?.length === 4
        ? { width: viewBox[2], height: viewBox[3] }
        : { width: WIDTH, height: HEIGHT }
    })
    expect(canvas).toEqual({
      width: visualBaselines.width,
      height: visualBaselines.height,
    })
    await new Promise((r) => setTimeout(r, 500))

    // Make all fragments visible
    await page.addStyleTag({
      content:
        '[data-bespoke-marp-fragment=inactive]{visibility:visible!important;opacity:1!important}',
    })

    // Detect bespoke vs static
    const isBespoke: boolean = await page.evaluate(() => {
      const svgs = document.querySelectorAll('svg[data-marpit-svg]')
      if (svgs.length < 2) return false
      return svgs[1].getBoundingClientRect().top < 100
    })

    for (const idx of slideIndices) {
      const pngPath = path.join(outDir, `html-slide-${idx}.png`)
      if (isBespoke) {
        await page.evaluate((n: number) => {
          window.location.hash = '#' + n
        }, idx)
        await new Promise((r) => setTimeout(r, 300))
        await page.screenshot({ path: pngPath })
      } else {
        // Static: clip to the SVG area
        const clip = await page.evaluate((n: number) => {
          const section = document.getElementById(String(n))
          if (!section) return null
          const svg = section.closest('svg')
          if (!svg) return null
          const r = svg.getBoundingClientRect()
          return {
            x: Math.round(r.left + window.scrollX),
            y: Math.round(r.top + window.scrollY),
            width: Math.round(r.width),
            height: Math.round(r.height),
          }
        }, idx)
        if (clip) {
          await page.screenshot({
            path: pngPath,
            clip,
            captureBeyondViewport: true,
          })
        } else {
          // Fallback: scrollTo approach
          await page.screenshot({
            path: pngPath,
            clip: {
              x: 0,
              y: (idx - 1) * HEIGHT,
              width: WIDTH,
              height: HEIGHT,
            },
            captureBeyondViewport: true,
          })
        }
      }
      results.set(idx, pngPath)
    }
  } finally {
    await browser.close()
  }
  return results
}

// --- Helper: export PPTX slides to PNG via PowerPoint COM ---
function exportPptxSlides(
  pptxPath: string,
  outDir: string,
  slideIndices: number[],
): Map<number, string> {
  const psScript = `
param([string]$PptxPath, [string]$OutDir, [string]$Indices)
$slideNums = $Indices -split ',' | ForEach-Object { [int]$_ }
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
try {
  $pres = $app.Presentations.Open($PptxPath, $true, $false, $false)
  foreach ($n in $slideNums) {
    if ($n -le $pres.Slides.Count) {
      $outPath = Join-Path $OutDir ("pptx-slide-" + $n + ".png")
      $pres.Slides($n).Export($outPath, "PNG", ${WIDTH}, ${HEIGHT})
    }
  }
  $pres.Close()
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
`

  const psFile = path.join(outDir, '_export.ps1')
  fs.writeFileSync(psFile, psScript, 'utf-8')

  const result = spawnSync(
    'powershell',
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      psFile,
      '-PptxPath',
      pptxPath,
      '-OutDir',
      outDir,
      '-Indices',
      slideIndices.join(','),
    ],
    { encoding: 'utf-8', timeout: 120000 },
  )

  if (result.status !== 0) {
    throw new Error(
      `PowerPoint export failed: ${result.stderr || result.stdout}`,
    )
  }

  const results = new Map<number, string>()
  for (const idx of slideIndices) {
    const p = path.join(outDir, `pptx-slide-${idx}.png`)
    if (fs.existsSync(p)) results.set(idx, p)
  }
  return results
}

// --- Helper: pixel diff ---
function pixelDiff(
  img1Path: string,
  img2Path: string,
): { diffPct: number; maxRegionPct: number; diffPath: string } {
  const pixelmatchMod = require('pixelmatch')
  const pixelmatch = pixelmatchMod.default ?? pixelmatchMod
  const { PNG } = require('pngjs')

  const img1 = PNG.sync.read(fs.readFileSync(img1Path))
  const img2 = PNG.sync.read(fs.readFileSync(img2Path))

  const { normalizePngPair } = require('./tools/visual-diff-metrics')
  const normalized = normalizePngPair(img1, img2)
  const w = img1.width
  const h = img1.height
  const diff = new PNG({ width: w, height: h })
  const data1 = normalized.reference.data
  const data2 = normalized.comparison.data

  const numDiff = pixelmatch(data1, data2, diff.data, w, h, {
    threshold: 0.12, // tolerant of anti-aliasing differences
  })
  const diffPct = numDiff / (w * h)
  const { computeRegionalDiff } = require('./tools/visual-diff-metrics')
  const maxRegionPct =
    computeRegionalDiff(data1, data2, w, h).maxRegionPct * 100

  const diffPath = img1Path.replace('html-slide', 'diff-slide')
  fs.writeFileSync(diffPath, PNG.sync.write(diff))

  return { diffPct, maxRegionPct, diffPath }
}

// ═══════════════════════════════════════════════════════════════════════════════
// Test Suite
// ═══════════════════════════════════════════════════════════════════════════════

const describeOrSkip = CAN_RUN ? describe : describe.skip

describeOrSkip(
  'Gate 3: Visual regression — all fixture slides must remain stable',
  () => {
    const outDir = path.join(PROJECT_ROOT, 'dist', 'visual-regression-test')
    const htmlPath = FIXTURE_HTML_PATH
    const pptxPath = path.join(outDir, 'test-output.pptx')

    let htmlScreenshots: Map<number, string>
    let pptxScreenshots: Map<number, string>

    beforeAll(async () => {
      fs.mkdirSync(outDir, { recursive: true })

      // Generate PPTX from the test HTML
      await generatePptx(htmlPath, pptxPath)

      // Determine actual slide indices for the code block slides
      // slides-ci.html uses section IDs 1-based; slides 87b and 90 are near the end
      // We need to find the actual slide numbers in the HTML
      const puppeteer = require('puppeteer-core')
      const browser = await puppeteer.launch({
        executablePath: CHROME_PATH,
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu'],
      })
      let totalSlides = 0
      try {
        const page = await browser.newPage()
        await page.setViewport({ width: WIDTH, height: HEIGHT })
        await page.goto(pathToFileURL(htmlPath).href, {
          waitUntil: 'networkidle0',
        })
        // Count total slides
        totalSlides = await page.evaluate(() => {
          let n = 0
          while (document.getElementById(String(n + 1))) n++
          return n
        })
      } finally {
        await browser.close()
      }

      expect(totalSlides).toBe(96)
      expect(VISUAL_BASELINES_PCT).toHaveLength(totalSlides)
      expect(MAX_REGION_BASELINES_PCT).toHaveLength(totalSlides)
      const testSlideIndices = Array.from(
        { length: totalSlides },
        (_, i) => i + 1,
      )

      // Screenshot HTML slides
      htmlScreenshots = await screenshotHtmlSlides(
        htmlPath,
        outDir,
        testSlideIndices,
      )

      // Export PPTX slides to PNG
      pptxScreenshots = exportPptxSlides(pptxPath, outDir, testSlideIndices)
    }, 600000) // 10 minutes for all 96 PowerPoint renders

    it('all fixture slides have no structural visual diff above 6%', () => {
      const failures: string[] = []
      const results: string[] = []

      for (const [idx, htmlPng] of htmlScreenshots) {
        const pptxPng = pptxScreenshots.get(idx)
        if (!pptxPng) {
          failures.push(`Slide ${idx}: PPTX screenshot missing`)
          continue
        }

        const { diffPct, maxRegionPct, diffPath } = pixelDiff(htmlPng, pptxPng)
        const diffPctValue = diffPct * 100
        const baseline = VISUAL_BASELINES_PCT[idx - 1]
        const maxRegionBaseline = MAX_REGION_BASELINES_PCT[idx - 1]
        const gate = evaluateVisualGate(
          {
            wholeSlidePct: diffPctValue,
            maxRegionPct,
            wholeSlideBaselinePct: baseline,
            maxRegionBaselinePct: maxRegionBaseline,
          },
          {
            absoluteFailPct: FAIL_THRESHOLD * 100,
            wholeSlideTolerancePct: BASELINE_TOLERANCE_PCT,
            maxRegionTolerancePct: MAX_REGION_TOLERANCE_PCT,
          },
        )
        results.push(
          `Slide ${idx}: ${diffPctValue.toFixed(2)}% diff` +
            ` (limit ${gate.wholeSlideLimitPct.toFixed(2)}%), ` +
            `max region ${maxRegionPct.toFixed(2)}% ` +
            `(limit ${gate.maxRegionLimitPct.toFixed(2)}%)` +
            `${gate.passed ? '' : ' [FAIL]'}`,
        )
        if (!gate.passed) {
          failures.push(
            `Slide ${idx}: ${diffPctValue.toFixed(2)}% diff ` +
              `(whole-slide limit: ${gate.wholeSlideLimitPct.toFixed(2)}%, ` +
              `max region: ${maxRegionPct.toFixed(2)}%, ` +
              `region limit: ${gate.maxRegionLimitPct.toFixed(2)}%) — ` +
              `${gate.failures.join('; ')} — see ${path.basename(diffPath)}`,
          )
        }
      }

      // Write results summary for inspection
      const summaryPath = path.join(outDir, 'gate3-results.txt')
      fs.writeFileSync(summaryPath, results.join('\n'), 'utf-8')

      if (failures.length > 0) {
        throw new Error(
          `Visual regression detected on ${failures.length} slide(s):\n` +
            failures.join('\n') +
            `\n\nAll results:\n` +
            results.join('\n') +
            `\n\nDiff images saved in: ${outDir}`,
        )
      }
    })

    it('PPTX was generated without errors', () => {
      expect(fs.existsSync(pptxPath)).toBe(true)
      const stats = fs.statSync(pptxPath)
      expect(stats.size).toBeGreaterThan(10000) // At least 10KB
    })
  },
)
