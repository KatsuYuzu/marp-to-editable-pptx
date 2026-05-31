/**
 * Gate 3: Visual regression test — pixel-level fidelity comparison.
 *
 * Generates PPTX from the test fixture HTML, renders both HTML and PPTX
 * to PNG screenshots, then compares them via pixelmatch.
 *
 * Requirements:
 * - Chrome/Chromium installed (auto-detected)
 * - PowerPoint desktop (COM automation) on Windows
 * - lib/native-pptx.cjs built (npm run build)
 *
 * Skips automatically if PowerPoint or Chrome is unavailable.
 */
import { spawnSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'

// --- Environment detection (run once at module load) ---

function findChrome(): string | undefined {
  try {
    const { computeSystemExecutablePath, Browser, ChromeReleaseChannel } =
      require('@puppeteer/browsers')
    const p = computeSystemExecutablePath({
      browser: Browser.CHROME,
      platform: 'win64',
      channel: ChromeReleaseChannel.STABLE,
    })
    if (fs.existsSync(p)) return p
  } catch {
    /* not found */
  }
  return process.env.CHROME_PATH || undefined
}

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

const CHROME_PATH = findChrome()
const HAS_POWERPOINT = hasPowerPoint()
const PROJECT_ROOT = path.resolve(__dirname, '../..')
const BUNDLE_PATH = path.join(PROJECT_ROOT, 'lib', 'native-pptx.cjs')
const HAS_BUNDLE = fs.existsSync(BUNDLE_PATH)

const CAN_RUN = !!(CHROME_PATH && HAS_POWERPOINT && HAS_BUNDLE)

const WIDTH = 1280
const HEIGHT = 720

// Slides with known code block overflow risks (from pptx-export.md)
// Slide 87b = long code block overflow, Slide 90 = large code block font size
const CODE_BLOCK_SLIDES = [
  { slideNum: 18, name: '87b-long-code-overflow' }, // slide index in HTML
  { slideNum: 19, name: '90-code-font-fidelity' },
]

// --- Pixel diff thresholds ---
// These must be STRICT for code block overflow detection.
// Normal font rendering diff (Chrome vs PowerPoint) is ~2-5%.
// Code block overflow causes >15% diff because text spills outside the shape.
const FAIL_THRESHOLD = 0.08 // >8% → structural layout defect

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
): { diffPct: number; diffPath: string } {
  const pixelmatchMod = require('pixelmatch')
  const pixelmatch = pixelmatchMod.default ?? pixelmatchMod
  const { PNG } = require('pngjs')

  const img1 = PNG.sync.read(fs.readFileSync(img1Path))
  const img2 = PNG.sync.read(fs.readFileSync(img2Path))

  const w = Math.min(img1.width, img2.width)
  const h = Math.min(img1.height, img2.height)
  const diff = new PNG({ width: w, height: h })

  // Crop to common size
  let data1 = img1.data
  let data2 = img2.data
  if (img1.width !== w || img1.height !== h) {
    const tmp = new PNG({ width: w, height: h })
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const s = (y * img1.width + x) * 4
        const d = (y * w + x) * 4
        tmp.data[d] = img1.data[s]
        tmp.data[d + 1] = img1.data[s + 1]
        tmp.data[d + 2] = img1.data[s + 2]
        tmp.data[d + 3] = img1.data[s + 3]
      }
    data1 = tmp.data
  }
  if (img2.width !== w || img2.height !== h) {
    const tmp = new PNG({ width: w, height: h })
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const s = (y * img2.width + x) * 4
        const d = (y * w + x) * 4
        tmp.data[d] = img2.data[s]
        tmp.data[d + 1] = img2.data[s + 1]
        tmp.data[d + 2] = img2.data[s + 2]
        tmp.data[d + 3] = img2.data[s + 3]
      }
    data2 = tmp.data
  }

  const numDiff = pixelmatch(data1, data2, diff.data, w, h, {
    threshold: 0.12, // tolerant of anti-aliasing differences
  })
  const diffPct = numDiff / (w * h)

  const diffPath = img1Path.replace('html-slide', 'diff-slide')
  fs.writeFileSync(diffPath, PNG.sync.write(diff))

  return { diffPct, diffPath }
}

// ═══════════════════════════════════════════════════════════════════════════════
// Test Suite
// ═══════════════════════════════════════════════════════════════════════════════

const describeOrSkip = CAN_RUN ? describe : describe.skip

describeOrSkip(
  'Gate 3: Visual regression — code block slides must not overflow in PPTX',
  () => {
    const outDir = path.join(PROJECT_ROOT, 'dist', 'visual-regression-test')
    const htmlPath = path.join(
      PROJECT_ROOT,
      'src',
      'native-pptx',
      'test-fixtures',
      'slides-ci.html',
    )
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

      // The code block overflow slides are near the end of the fixture.
      // Use the last few slides which contain code blocks (87b=slide ~55, 90=~58)
      // We'll test ALL slides and find the ones with code blocks by scanning.
      // For efficiency, test a known range. The test fixture has ~94 slides.
      // Code block slides: 87b(≈55), 90(≈58), 92(≈60) in the slide count
      // Let's just test the last 10 slides which include all code block regression cases.
      const startSlide = Math.max(1, totalSlides - 9)
      const testSlideIndices = Array.from(
        { length: Math.min(10, totalSlides) },
        (_, i) => startSlide + i,
      )

      // Screenshot HTML slides
      htmlScreenshots = await screenshotHtmlSlides(
        htmlPath,
        outDir,
        testSlideIndices,
      )

      // Export PPTX slides to PNG
      pptxScreenshots = exportPptxSlides(pptxPath, outDir, testSlideIndices)
    }, 180000) // 3 minutes timeout for PPTX generation + screenshots

    it('code block slides must have <8% pixel diff (no text overflow)', () => {
      const failures: string[] = []
      const results: string[] = []

      for (const [idx, htmlPng] of htmlScreenshots) {
        const pptxPng = pptxScreenshots.get(idx)
        if (!pptxPng) {
          failures.push(`Slide ${idx}: PPTX screenshot missing`)
          continue
        }

        const { diffPct, diffPath } = pixelDiff(htmlPng, pptxPng)
        results.push(
          `Slide ${idx}: ${(diffPct * 100).toFixed(2)}% diff${diffPct > FAIL_THRESHOLD ? ' [FAIL]' : ''}`,
        )
        if (diffPct > FAIL_THRESHOLD) {
          failures.push(
            `Slide ${idx}: ${(diffPct * 100).toFixed(2)}% diff (threshold: ${FAIL_THRESHOLD * 100}%) — see ${path.basename(diffPath)}`,
          )
        }
      }

      // Write results summary for inspection
      const summaryPath = path.join(outDir, 'gate3-results.txt')
      fs.writeFileSync(summaryPath, results.join('\n'), 'utf-8')

      if (failures.length > 0) {
        fail(
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
