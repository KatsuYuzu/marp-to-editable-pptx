#!/usr/bin/env node
/**
 * Batch structural test runner.
 * Runs structural-test.js on all PPTX files in dist/vr-batch/.
 *
 * Usage: node scripts/run-structural-batch.js
 */
const { spawnSync } = require('child_process')
const fs = require('fs')
const path = require('path')

const root = path.resolve(__dirname, '..')
process.chdir(root)

const batchDir = path.join(root, 'dist', 'vr-batch')
if (!fs.existsSync(batchDir)) {
  console.error('dist/vr-batch not found. Run VR batch test first.')
  process.exit(1)
}

const htmlFiles = fs.readdirSync(batchDir)
  .filter(f => f.endsWith('.html'))
  .sort()

console.log(`Found ${htmlFiles.length} HTML files in dist/vr-batch\n`)

const results = []
let totalFail = 0, totalWarn = 0, totalOk = 0, totalSlides = 0

for (let i = 0; i < htmlFiles.length; i++) {
  const htmlFile = htmlFiles[i]
  const baseName = htmlFile.replace(/\.html$/, '')
  const htmlPath = path.join(batchDir, htmlFile)
  const pptxPath = path.join(batchDir, baseName + '.pptx')

  if (!fs.existsSync(pptxPath)) {
    console.log(`[${i + 1}/${htmlFiles.length}] ${baseName} — SKIP (no PPTX)`)
    continue
  }

  const pngDir = path.join(root, 'dist', `compare-${baseName}`)

  process.stdout.write(`[${i + 1}/${htmlFiles.length}] ${baseName}...`)

  const args = [
    path.join(root, 'src', 'native-pptx', 'tools', 'structural-test.js'),
    htmlPath,
    pptxPath,
  ]
  if (fs.existsSync(pngDir)) {
    args.push(`--png-dir=${pngDir}`)
  }

  const result = spawnSync('node', args, {
    encoding: 'utf-8',
    timeout: 180000,
    stdio: ['pipe', 'pipe', 'pipe'],
  })

  // Parse JSON from stdout
  const stdout = result.stdout || ''
  try {
    const report = JSON.parse(stdout)
    const f = report.summary.FAIL
    const w = report.summary.WARN
    const o = report.summary.OK
    totalFail += f
    totalWarn += w
    totalOk += o
    totalSlides += report.slides

    const failInfo = report.perSlide
      .filter(s => s.overall === 'FAIL')
      .map(s => {
        const reasons = []
        if (s.content !== 'OK') reasons.push('C:' + s.details.contentMissingPct + '%')
        if (s.structure !== 'OK') reasons.push('S:' + s.details.structureIssues.join(','))
        if (s.tofu !== 'OK') reasons.push('T')
        return `${s.slide}(${reasons.join('+')})`
      })

    const mark = f > 0 ? '✗' : w > 0 ? '△' : '○'
    console.log(` ${mark} F${f}/W${w}/OK${o}${failInfo.length > 0 ? ' FAIL:' + failInfo.join(',') : ''}`)
    results.push({ file: baseName, ...report.summary, slides: report.slides, failedSlides: failInfo })
  } catch (e) {
    console.log(` ERROR (parse failed)`)
    if (result.stderr) console.log(`  stderr: ${result.stderr.slice(0, 200)}`)
    results.push({ file: baseName, FAIL: 0, WARN: 0, OK: 0, slides: 0, error: true })
  }
}

// Summary
console.log('\n' + '='.repeat(60))
console.log('STRUCTURAL TEST BATCH SUMMARY')
console.log('='.repeat(60))
console.log(`Files: ${results.length}`)
console.log(`Slides: ${totalSlides} — FAIL:${totalFail} WARN:${totalWarn} OK:${totalOk}`)
if (totalSlides > 0) {
  console.log(`  FAIL rate: ${((totalFail / totalSlides) * 100).toFixed(1)}%`)
  console.log(`  WARN rate: ${((totalWarn / totalSlides) * 100).toFixed(1)}%`)
  console.log(`  OK   rate: ${((totalOk / totalSlides) * 100).toFixed(1)}%`)
}

console.log('\nFailed files:')
for (const r of results) {
  if (r.FAIL > 0) {
    console.log(`  ✗ ${r.file}: F${r.FAIL}/W${r.WARN}/OK${r.OK} [${r.failedSlides.join(', ')}]`)
  }
}

// Write JSON report
const reportPath = path.join(batchDir, 'structural-report.json')
fs.writeFileSync(reportPath, JSON.stringify(results, null, 2))
console.log(`\nReport: ${reportPath}`)
