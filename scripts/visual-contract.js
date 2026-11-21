#!/usr/bin/env node
const fs = require('node:fs')
const path = require('node:path')
const { PNG } = require('pngjs')
const {
  computeRegionalDiff,
} = require('../src/native-pptx/tools/visual-diff-metrics')

function compareDirectories(
  htmlDir,
  pptxDir,
  {
    failThreshold = 0.12,
    regionFailThreshold = 0.7,
    columns = 16,
    rows = 9,
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

  return htmlFiles.map((htmlName, index) => {
    const html = PNG.sync.read(fs.readFileSync(path.join(htmlDir, htmlName)))
    const pptx = PNG.sync.read(
      fs.readFileSync(path.join(pptxDir, pptxFiles[index])),
    )
    const width = Math.min(html.width, pptx.width)
    const height = Math.min(html.height, pptx.height)
    const crop = (image) => {
      if (image.width === width && image.height === height) return image.data
      const data = new Uint8Array(width * height * 4)
      for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
          const source = (y * image.width + x) * 4
          const target = (y * width + x) * 4
          data.set(image.data.subarray(source, source + 4), target)
        }
      }
      return data
    }
    const metrics = computeRegionalDiff(
      crop(html),
      crop(pptx),
      width,
      height,
      { columns, rows },
    )
    return {
      slide: index + 1,
      wholeSlidePct: metrics.wholeSlidePct,
      maxRegionPct: metrics.maxRegionPct,
      status:
        metrics.wholeSlidePct > failThreshold ||
        metrics.maxRegionPct > regionFailThreshold
          ? 'FAIL'
          : 'OK',
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
    const results = compareDirectories(htmlDir, pptxDir)
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
