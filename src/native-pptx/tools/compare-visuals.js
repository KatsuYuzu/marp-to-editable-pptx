#!/usr/bin/env node
/**
 * Visual fidelity comparison tool for native PPTX.
 *
 * Usage:
 *   node src/native-pptx/tools/compare-visuals.js <marp-html-path> <pptx-path> [chrome-path]
 *
 * Outputs (in a folder next to the HTML):
 *   html-slide-NNN.png  — screenshot of each Marp HTML slide
 *   pptx-slide-NNN.png  — screenshot via PPTX → PowerPoint COM → PNG
 *   compare-NNN.png     — side-by-side diff image (HTML left | PPTX right)
 *   compare-report.md   — textual summary of diff areas
 */
const { spawnSync } = require('node:child_process')
const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')

/** Auto-detect Chrome/Chromium on the host system. */
function findChrome() {
  const {
    computeSystemExecutablePath,
    Browser,
    ChromeReleaseChannel,
  } = require('@puppeteer/browsers')
  const platforms = { win32: 'win64', darwin: 'mac', linux: 'linux' }
  const platform = platforms[process.platform]
  if (!platform) return undefined
  try {
    const p = computeSystemExecutablePath({
      browser: Browser.CHROME,
      platform,
      channel: ChromeReleaseChannel.STABLE,
    })
    if (fs.existsSync(p)) return p
  } catch {
    /* not found */
  }
  return undefined
}

const WIDTH = 1280
const HEIGHT = 720

async function main() {
  const htmlPath = path.resolve(process.argv[2])
  const pptxPath = path.resolve(process.argv[3])
  const chromePath = process.argv[4] ?? process.env.CHROME_PATH ?? findChrome()

  if (!fs.existsSync(htmlPath)) {
    console.error('HTML not found:', htmlPath)
    process.exit(1)
  }
  if (!fs.existsSync(pptxPath)) {
    console.error('PPTX not found:', pptxPath)
    process.exit(1)
  }

  // Output always goes to dist/ under the project root so that generated
  // comparison artifacts never land inside the source tree.
  // compare-visuals.js lives 3 levels deep (src/native-pptx/tools/), so three
  // path.resolve levels up reaches the project root.
  const projectRoot = path.resolve(__dirname, '../../..')
  const outDir = path.join(
    projectRoot,
    'dist',
    'compare-' + path.basename(htmlPath, '.html'),
  )
  fs.mkdirSync(outDir, { recursive: true })
  console.log('Output dir:', outDir)

  // ─── Step 1: HTML slide screenshots via Puppeteer ───────────────────────
  const puppeteer = require('puppeteer-core')
  const browser = await puppeteer.launch({
    executablePath: chromePath,
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
    ],
  })

  try {
    const page = await browser.newPage()
    await page.setViewport({ width: WIDTH, height: HEIGHT })
    await page.goto(pathToFileURL(htmlPath).href, { waitUntil: 'networkidle0' })

    // Let bespoke.js finish initializing
    await new Promise((r) => setTimeout(r, 500))

    // Force all bespoke fragments visible so screenshots show full slide content
    await page.addStyleTag({
      content:
        '[data-bespoke-marp-fragment=inactive]{visibility:visible!important;opacity:1!important}',
    })
    // Hide bespoke On-Screen Controller (navigation arrows/buttons) so they
    // never appear in screenshots. No-op for static non-bespoke HTMLs.
    await page.addStyleTag({
      content:
        '[data-bespoke-marp-osc]{display:none!important}.bespoke-marp-osc{display:none!important}',
    })

    // Count slides via accurate key-based grouping (matches the PPTX extractor).
    //
    // Marp's advanced-background feature generates multiple <section> layers
    // per slide (background / content / pseudo).  A simple DOM query for
    // sections *without* the attribute misses "content"-layer sections, giving
    // a lower count than the actual number of exportable slides.
    //
    // `window.bespoke.slides.length` is similarly unreliable: it reflects the
    // bespoke navigation count which may exclude advanced-background content
    // sections.  We therefore always use the key-based count.
    const slideCount = await page.evaluate(() => {
      const allSections = Array.from(document.querySelectorAll('section'))
        .filter((s) => {
          if (s.parentElement?.closest('section')) return false
          return (
            s.parentElement?.tagName.toLowerCase() === 'foreignobject' ||
            s.hasAttribute('data-marpit-pagination')
          )
        })
      const keys = new Set()
      allSections.forEach((s, i) => {
        const layer = s.getAttribute('data-marpit-advanced-background')
        if (layer === 'pseudo') return
        const key =
          s.getAttribute('data-marpit-pagination') ||
          s.getAttribute('id') ||
          String(i)
        keys.add(key)
      })
      return keys.size
    })

    console.log(`HTML slide count: ${slideCount}`)

    // Detect whether the HTML is bespoke.js-powered or static.
    //
    // Bespoke HTML (default marp CLI output): all slide SVGs are absolutely/
    // fixed-positioned at (0,0) and overlap — only the active slide is visible.
    // window.bespoke is NOT exposed as a global in modern marp-cli bundles, so
    // we detect bespoke by checking the layout: if the second SVG is at top=0
    // (same as first), the slides are stacked → bespoke mode.
    //
    // Static HTML (marp.render() output): SVGs flow vertically as block
    // elements — the second SVG is at top≈720. Hash navigation has no effect.
    const isBespoke = await page.evaluate(() => {
      const svgs = document.querySelectorAll('svg[data-marpit-svg]')
      if (svgs.length < 2) return false
      return svgs[1].getBoundingClientRect().top < 100
    })

    if (isBespoke) {
      // ── Bespoke HTML: hash navigation per slide ──────────────────────────
      // Marp's bespoke hash uses 1-based indexing: #1 = slide 1, #2 = slide 2.
      // window.location.hash change triggers bespoke to activate the target
      // slide and position it in the viewport at (0,0).
      for (let i = 0; i < slideCount; i++) {
        await page.evaluate((n) => {
          window.location.hash = '#' + n
        }, i + 1)
        // Wait for bespoke to complete the slide transition.
        await new Promise((r) => setTimeout(r, 200))
        const slidePng = path.join(
          outDir,
          `html-slide-${String(i + 1).padStart(3, '0')}.png`,
        )
        await page.screenshot({ path: slidePng })
        process.stdout.write(`  HTML slide ${i + 1}/${slideCount} saved\r`)
      }
    } else {
      // ── Static HTML: SVG clip approach ───────────────────────────────────
      //
      // Static Marp HTML from marp.render() places each slide in its own
      // <svg data-marpit-svg> element.  We locate each slide's SVG via
      // its contained <section id="N"> and clip the screenshot to the SVG's
      // page-coordinate bounding rect.
      //
      // For advanced-background slides (![bg]), Marp emits three SVG layers
      // (background, content, pseudo).  The "content" section keeps the numeric
      // id; its parent SVG is clipped for the screenshot.
      const slideClips = await page.evaluate(() => {
        const clips = []
        for (let n = 1; ; n++) {
          const section = document.getElementById(String(n))
          if (!section) break
          const svg = section.closest('svg')
          if (!svg) {
            // The section exists but has no SVG ancestor (e.g. advanced background
            // layout). Record null so the index stays aligned with slide numbers.
            clips.push(null)
            continue
          }
          const r = svg.getBoundingClientRect()
          clips.push({
            x: Math.round(r.left + window.scrollX),
            y: Math.round(r.top + window.scrollY),
            width: Math.round(r.width),
            height: Math.round(r.height),
          })
        }
        return clips
      })

      for (let i = 0; i < slideCount; i++) {
        const clip =
          slideClips[i] ?? { x: 0, y: i * HEIGHT, width: WIDTH, height: HEIGHT }
        const slidePng = path.join(
          outDir,
          `html-slide-${String(i + 1).padStart(3, '0')}.png`,
        )
        await page.screenshot({
          path: slidePng,
          clip,
          captureBeyondViewport: true,
        })
        process.stdout.write(`  HTML slide ${i + 1}/${slideCount} saved\r`)
      }
    }
    console.log('\n  HTML slides done.')
  } finally {
    await browser.close()
  }

  // ─── Step 2: PPTX slide screenshots via PowerPoint COM ──────────────────
  console.log('Exporting PPTX slides via PowerShell/PowerPoint COM...')

  // Write a PS1 that takes paths via parameters to avoid escaping issues
  const psScriptPath = path.join(outDir, '_export-pptx.ps1')
  const psScript = `param(
  [string]$PptxPath,
  [string]$OutDir,
  [int]$Width = ${WIDTH},
  [int]$Height = ${HEIGHT}
)
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
try {
  $pres = $app.Presentations.Open($PptxPath, $true, $false, $false)
  $slideCount = $pres.Slides.Count
  Write-Host "PPTX slide count: $slideCount"
  for ($i = 1; $i -le $slideCount; $i++) {
    $slide = $pres.Slides($i)
    $outPath = Join-Path $OutDir ("pptx-slide-" + $i.ToString("D3") + ".png")
    $slide.Export($outPath, "PNG", $Width, $Height)
    Write-Host ("  PPTX slide $i/$slideCount saved")
  }
  $pres.Close()
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  [System.GC]::Collect()
}
`
  fs.writeFileSync(psScriptPath, psScript, 'utf-8')

  const psResult = spawnSync(
    'powershell',
    [
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      psScriptPath,
      '-PptxPath',
      pptxPath,
      '-OutDir',
      outDir,
    ],
    {
      encoding: 'utf-8',
      timeout: 300000, // 5 minutes (increased from 2 min for large decks)
    },
  )
  if (psResult.stdout) console.log(psResult.stdout)
  if (psResult.stderr) console.error('  PS STDERR:', psResult.stderr)

  // ─── Step 3: Pixel diff + Side-by-side comparison HTML report ──────────
  const pixelmatchMod = require('pixelmatch')
  const pixelmatch = pixelmatchMod.default ?? pixelmatchMod
  const { PNG } = require('pngjs')

  /**
   * Threshold for classifying a slide as failed/warned (fraction of pixels).
   *
   * HTML → PPTX conversion cannot be pixel-perfect: Chrome and PowerPoint use
   * different font rendering engines (Blink vs GDI+), causing heading heights
   * and line-break positions to differ by a few pixels.  These differences
   * cascade vertically, producing ~5-7% diff on text-heavy slides even when
   * the content is entirely correct.  7.5% is calibrated to the observed
   * font-metric noise floor; anything above it indicates a real layout defect.
   */
  const FAIL_THRESHOLD = 0.075 // >7.5% different pixels → FAIL (content defect)
  const WARN_THRESHOLD = 0.01 // >1% → WARN (font rendering noise)

  const htmlSlides = fs
    .readdirSync(outDir)
    .filter((f) => f.startsWith('html-slide-'))
    .sort()
  const pptxSlides = fs
    .readdirSync(outDir)
    .filter((f) => f.startsWith('pptx-slide-'))
    .sort()

  const maxSlides = Math.max(htmlSlides.length, pptxSlides.length)
  if (maxSlides === 0) {
    console.error('No slides found in output dir')
    process.exit(1)
  }

  console.log('\nRunning pixel diff...')

  /** @type {{ n: number, status: 'FAIL'|'WARN'|'OK'|'MISSING', diffPct: number }[]} */
  const slideResults = []

  for (let i = 0; i < maxSlides; i++) {
    const n = i + 1
    const pad = String(n).padStart(3, '0')
    const htmlImgPath = path.join(outDir, `html-slide-${pad}.png`)
    const pptxImgPath = path.join(outDir, `pptx-slide-${pad}.png`)
    const diffImgPath = path.join(outDir, `diff-slide-${pad}.png`)

    if (!fs.existsSync(htmlImgPath) || !fs.existsSync(pptxImgPath)) {
      slideResults.push({ n, status: 'MISSING', diffPct: 1 })
      continue
    }

    try {
      const img1 = PNG.sync.read(fs.readFileSync(htmlImgPath))
      const img2 = PNG.sync.read(fs.readFileSync(pptxImgPath))

      // Ensure both images are the same size (use the smaller dimensions)
      const w = Math.min(img1.width, img2.width)
      const h = Math.min(img1.height, img2.height)
      const diff = new PNG({ width: w, height: h })

      // Crop to common size if needed
      let data1 = img1.data,
        data2 = img2.data
      if (img1.width !== w || img1.height !== h) {
        // Reallocate with correct size
        const tmp = new PNG({ width: w, height: h })
        for (let y = 0; y < h; y++)
          for (let x = 0; x < w; x++) {
            const s = (y * img1.width + x) * 4,
              d = (y * w + x) * 4
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
            const s = (y * img2.width + x) * 4,
              d = (y * w + x) * 4
            tmp.data[d] = img2.data[s]
            tmp.data[d + 1] = img2.data[s + 1]
            tmp.data[d + 2] = img2.data[s + 2]
            tmp.data[d + 3] = img2.data[s + 3]
          }
        data2 = tmp.data
      }

      const numDiff = pixelmatch(data1, data2, diff.data, w, h, {
        threshold: 0.1,
      })
      const diffPct = numDiff / (w * h)
      const status =
        diffPct > FAIL_THRESHOLD
          ? 'FAIL'
          : diffPct > WARN_THRESHOLD
            ? 'WARN'
            : 'OK'

      fs.writeFileSync(diffImgPath, PNG.sync.write(diff))
      slideResults.push({ n, status, diffPct })
      process.stdout.write(
        `  Slide ${pad}: ${status} (${(diffPct * 100).toFixed(2)}% diff)\n`,
      )
    } catch (e) {
      slideResults.push({ n, status: 'FAIL', diffPct: 1 })
      console.warn(`  Slide ${pad}: pixel diff failed — ${e.message}`)
    }
  }

  // Print summary
  const fails = slideResults.filter((r) => r.status === 'FAIL')
  const warns = slideResults.filter((r) => r.status === 'WARN')
  const oks = slideResults.filter((r) => r.status === 'OK')
  console.log(`\n=== DIFF SUMMARY ===`)
  console.log(
    `  FAIL: ${fails.length}  WARN: ${warns.length}  OK: ${oks.length}  MISSING: ${slideResults.filter((r) => r.status === 'MISSING').length}`,
  )
  if (fails.length > 0)
    console.log(`  FAILed slides: ${fails.map((r) => r.n).join(', ')}`)
  if (warns.length > 0)
    console.log(`  WARNed slides: ${warns.map((r) => r.n).join(', ')}`)

  // Generate an HTML comparison report
  const STATUS_COLOR = {
    FAIL: '#f44',
    WARN: '#fa0',
    OK: '#4c4',
    MISSING: '#aaa',
  }
  const rows = []
  for (let i = 0; i < maxSlides; i++) {
    const n = i + 1
    const pad = String(n).padStart(3, '0')
    const htmlImg = htmlSlides[i] ? `html-slide-${pad}.png` : null
    const pptxImg = pptxSlides[i] ? `pptx-slide-${pad}.png` : null
    const diffImg = fs.existsSync(path.join(outDir, `diff-slide-${pad}.png`))
      ? `diff-slide-${pad}.png`
      : null
    const result = slideResults.find((r) => r.n === n) ?? {
      status: 'MISSING',
      diffPct: 1,
    }
    const bgColor = STATUS_COLOR[result.status]
    const diffLabel =
      result.status === 'MISSING'
        ? 'MISSING'
        : `${result.status} (${(result.diffPct * 100).toFixed(2)}%)`
    rows.push(`
    <tr>
      <td style="padding:4px;font-weight:bold;vertical-align:middle;background:${bgColor};color:#fff;text-align:center">
        ${n}<br><small>${diffLabel}</small>
      </td>
      <td style="padding:4px">
        ${htmlImg ? `<img src="${htmlImg}" width="426" style="border:1px solid #ccc" alt="HTML ${n}">` : '<em>missing</em>'}
        <br><small>Marp HTML</small>
      </td>
      <td style="padding:4px">
        ${pptxImg ? `<img src="${pptxImg}" width="426" style="border:1px solid #ccc" alt="PPTX ${n}">` : '<em>missing</em>'}
        <br><small>PPTX (native)</small>
      </td>
      <td style="padding:4px">
        ${diffImg ? `<img src="${diffImg}" width="426" style="border:1px solid #ccc" alt="DIFF ${n}">` : '<em>—</em>'}
        <br><small>Pixel diff</small>
      </td>
    </tr>`)
  }

  // Escape HTML special characters to prevent broken markup or XSS when a
  // crafted file name contains '<', '>', '&', or '"' characters.
  const escHtml = (s) =>
    s
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
  const escapedHtmlName = escHtml(path.basename(htmlPath))

  const reportHtml = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Native PPTX Fidelity Comparison — ${escapedHtmlName}</title>
<style>
  body { font-family: sans-serif; margin: 16px; background: #f5f5f5; }
  h1 { font-size: 18px; }
  .summary { background:#fff; border:1px solid #ccc; padding:8px 16px; margin-bottom:12px; border-radius:4px; }
  .badge { display:inline-block; padding:2px 8px; border-radius:3px; color:#fff; font-weight:bold; margin:2px; }
  .FAIL  { background: #f44; }
  .WARN  { background: #fa0; }
  .OK    { background: #4c4; }
  .MISSING { background: #aaa; }
  table { border-collapse: collapse; width: 100%; }
  th { background: #333; color: #fff; padding: 6px 12px; text-align: left; }
  tr:nth-child(even) { background: #eee; }
  td { vertical-align: top; }
  img { display: block; }
</style>
</head>
<body>
<h1>Native PPTX Fidelity: ${escapedHtmlName}</h1>
<p>Generated: ${new Date().toLocaleString('ja-JP')}</p>
<div class="summary">
  <span class="badge FAIL">FAIL ${fails.length}</span>
  <span class="badge WARN">WARN ${warns.length}</span>
  <span class="badge OK">OK ${oks.length}</span>
  ${fails.length > 0 ? `<br><strong>Failed slides:</strong> ${fails.map((r) => `<a href="#slide-${r.n}">${r.n}</a>`).join(', ')}` : ''}
  ${warns.length > 0 ? `<br><strong>Warning slides:</strong> ${warns.map((r) => `<a href="#slide-${r.n}">${r.n}</a>`).join(', ')}` : ''}
</div>
<table>
  <thead>
    <tr>
      <th style="width:80px"># / Score</th>
      <th>Source (Marp HTML)</th>
      <th>Output (Native PPTX)</th>
      <th>Pixel Diff</th>
    </tr>
  </thead>
  <tbody>
    ${rows.join('\n')}
  </tbody>
</table>
</body>
</html>`

  const reportPath = path.join(outDir, 'compare-report.html')
  fs.writeFileSync(reportPath, reportHtml, 'utf-8')
  console.log('\nComparison report:', reportPath)
  console.log(
    `  ${maxSlides} slides compared (HTML: ${htmlSlides.length}, PPTX: ${pptxSlides.length})`,
  )

  // Exit with non-zero code if any FAIL so CI/loops can detect regressions
  if (fails.length > 0) process.exit(1)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
