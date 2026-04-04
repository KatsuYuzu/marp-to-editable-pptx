#!/usr/bin/env node
/**
 * Generate HTML and PPTX slide screenshots for CI.
 *
 * This script is a Linux-compatible equivalent of compare-visuals.js.
 * Instead of PowerPoint COM (Windows-only), it accepts pre-rendered PPTX
 * PNGs produced by LibreOffice (handled by the GitHub Actions workflow).
 *
 * Usage:
 *   node scripts/gen-html-screenshots.js <html-path> <output-dir> [chrome-path]
 *
 * Output (in <output-dir>):
 *   html-slide-001.png, html-slide-002.png, ...
 */
const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')

const WIDTH = 1280
const HEIGHT = 720

async function main() {
  const htmlPath = path.resolve(process.argv[2])
  const outDir = path.resolve(process.argv[3])
  const chromePath =
    process.argv[4] ?? process.env.CHROME_PATH ?? findChrome()

  if (!fs.existsSync(htmlPath)) {
    console.error('HTML file not found:', htmlPath)
    process.exit(1)
  }
  if (!chromePath) {
    console.error(
      'Chrome not found. Set CHROME_PATH env var or pass as 3rd argument.',
    )
    process.exit(1)
  }

  fs.mkdirSync(outDir, { recursive: true })

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

    // Force all bespoke fragments visible
    await page.addStyleTag({
      content:
        '[data-bespoke-marp-fragment=inactive]{visibility:visible!important;opacity:1!important}',
    })
    // Hide bespoke On-Screen Controller
    await page.addStyleTag({
      content:
        '[data-bespoke-marp-osc]{display:none!important}.bespoke-marp-osc{display:none!important}',
    })

    const slideCount = await page.evaluate(() => {
      // Use the maximum data-marpit-pagination value as the authoritative slide
      // count. window.bespoke.slides may return fewer slides than are actually
      // present (e.g. when advanced-background sections skew the DOM query),
      // and the SVG section query can also miscount due to attribute variations.
      const allSections = document.querySelectorAll('section[data-marpit-pagination]')
      let max = 0
      allSections.forEach((el) => {
        const v = parseInt(el.getAttribute('data-marpit-pagination') || '0', 10)
        if (v > max) max = v
      })
      if (max > 0) return max
      // Fallback: bespoke slide count
      if (window.bespoke?.slides) return window.bespoke.slides.length
      return 0
    })

    console.log(`HTML slide count: ${slideCount}`)

    for (let i = 0; i < slideCount; i++) {
      await page.evaluate((n) => {
        window.location.hash = '#' + n
      }, i + 1)
      await new Promise((r) => setTimeout(r, 300))

      const slidePng = path.join(
        outDir,
        `html-slide-${String(i + 1).padStart(3, '0')}.png`,
      )
      await page.screenshot({
        path: slidePng,
        clip: { x: 0, y: 0, width: WIDTH, height: HEIGHT },
      })
      console.log(`  Saved: ${path.basename(slidePng)}`)
    }
  } finally {
    await browser.close()
  }

  console.log('Done.')
}

function findChrome() {
  // Common paths on Linux CI (GitHub Actions ubuntu-latest with google-chrome)
  const candidates = [
    process.env.CHROME_PATH,
    '/usr/bin/google-chrome-stable',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium-browser',
    '/usr/bin/chromium',
  ].filter(Boolean)

  for (const p of candidates) {
    if (fs.existsSync(p)) return p
  }
  return undefined
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
