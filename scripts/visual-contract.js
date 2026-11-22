#!/usr/bin/env node
const fs = require('node:fs')
const path = require('node:path')
const { PNG } = require('pngjs')
const {
  computeRegionalDiff,
  normalizePngPair,
} = require('../src/native-pptx/tools/visual-diff-metrics')

function compareDirectories(
  htmlDir,
  pptxDir,
  {
    failThreshold = 0.12,
    regionFailThreshold = 0.7,
    columns = 16,
    rows = 9,
    baseline,
  } = {},
) {
  const htmlFiles = fs
    .readdirSync(htmlDir)
    .filter((name) => /^html-slide-\d+\.png$/.test(name))
    .sort()
  const pptxFiles = fs
    .readdirSync(pptxDir)
    .filter((name) => /^pptx-slide-\d+\.png$/.test(name))
    .sort()
  if (htmlFiles.length !== pptxFiles.length) {
    throw new Error(
      `Slide count mismatch: HTML=${htmlFiles.length} PPTX=${pptxFiles.length}`,
    )
  }
  if (baseline && baseline.slides.length !== htmlFiles.length) {
    throw new Error(
      `Baseline count mismatch: baseline=${baseline.slides.length} ` +
        `HTML=${htmlFiles.length}`,
    )
  }

  return htmlFiles.map((htmlName, index) => {
    const html = PNG.sync.read(fs.readFileSync(path.join(htmlDir, htmlName)))
    const pptx = PNG.sync.read(
      fs.readFileSync(path.join(pptxDir, pptxFiles[index])),
    )
    const normalized = normalizePngPair(html, pptx)
    const width = html.width
    const height = html.height
    const metrics = computeRegionalDiff(
      normalized.reference.data,
      normalized.comparison.data,
      width,
      height,
      { columns, rows },
    )
    const wholeSlideBaselinePct = baseline?.slides[index]?.wholeSlidePct
    const maxRegionBaselinePct = baseline?.slides[index]?.maxRegionPct
    const baselineFailure = baseline
      ? metrics.wholeSlidePct * 100 >
          wholeSlideBaselinePct + baseline.wholeSlideTolerancePct ||
        metrics.maxRegionPct * 100 >
          maxRegionBaselinePct + baseline.maxRegionTolerancePct
      : false
    const failed = baseline
      ? baselineFailure
      : metrics.wholeSlidePct > failThreshold ||
        metrics.maxRegionPct > regionFailThreshold
    return {
      slide: index + 1,
      width,
      height,
      wholeSlidePct: metrics.wholeSlidePct,
      maxRegionPct: metrics.maxRegionPct,
      status: failed ? 'FAIL' : 'OK',
    }
  })
}

if (require.main === module) {
  const [htmlDir, pptxDir, outputPath] = process.argv.slice(2)
  if (!htmlDir || !pptxDir) {
    console.error(
      'Usage: node scripts/visual-contract.js <html-dir> <pptx-dir> [output.json]',
    )
    process.exit(2)
  }
  try {
    const baselinePath = process.env.MARP_LIBREOFFICE_VISUAL_BASELINE
      ? path.resolve(process.env.MARP_LIBREOFFICE_VISUAL_BASELINE)
      : path.resolve(
          __dirname,
          '..',
          'src',
          'native-pptx',
          'test-fixtures',
          'libreoffice-visual-baselines.json',
        )
    const baseline = fs.existsSync(baselinePath)
      ? JSON.parse(fs.readFileSync(baselinePath, 'utf8'))
      : undefined
    const results = compareDirectories(htmlDir, pptxDir, { baseline })
    const payload = {
      renderer: 'libreoffice',
      slides: results,
      failures: results.filter((result) => result.status === 'FAIL'),
    }
    if (outputPath) {
      fs.writeFileSync(outputPath, JSON.stringify(payload, null, 2))
    }
    console.log(
      `Visual contract: ${payload.failures.length} FAIL, ` +
        `${results.length - payload.failures.length} OK`,
    )
    for (const result of results) {
      console.log(
        `  Slide ${String(result.slide).padStart(3, '0')}: ` +
          `${result.status} (${(result.wholeSlidePct * 100).toFixed(2)}%, ` +
          `max region ${(result.maxRegionPct * 100).toFixed(2)}%)`,
      )
    }
    if (payload.failures.length > 0) process.exit(1)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
}

module.exports = { compareDirectories }
