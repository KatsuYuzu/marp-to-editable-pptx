#!/usr/bin/env node
/**
 * Batch VR test: run compare-visuals on all .md files in a given folder.
 * Usage: node scripts/run-vr-test-batch.js <folder-path>
 */
const { spawnSync } = require('child_process')
const fs = require('fs')
const path = require('path')

const root = path.resolve(__dirname, '..')
process.chdir(root)

const inputDir = path.resolve(process.argv[2] || 'C:\\Users\\k.shimizu\\Downloads\\test')
if (!fs.existsSync(inputDir)) {
  console.error('Input directory not found:', inputDir)
  process.exit(1)
}

const mdFiles = fs.readdirSync(inputDir)
  .filter(f => f.endsWith('.md'))
  .sort()

console.log(`Found ${mdFiles.length} .md files in ${inputDir}\n`)

const distDir = path.join(root, 'dist', 'vr-batch')
fs.mkdirSync(distDir, { recursive: true })

const results = []

for (let i = 0; i < mdFiles.length; i++) {
  const mdFile = mdFiles[i]
  const baseName = mdFile.replace(/\.md$/, '')
  const mdPath = path.join(inputDir, mdFile)
  const htmlPath = path.join(distDir, baseName + '.html')
  const pptxPath = path.join(distDir, baseName + '.pptx')
  const compareDir = path.join(distDir, baseName)

  console.log(`[${i + 1}/${mdFiles.length}] ${mdFile}`)

  // Step 1: md -> html
  const mdToHtml = spawnSync('node', [
    path.join(root, 'src', 'native-pptx', 'tools', 'md-to-html.js'),
    mdPath,
    htmlPath
  ], { encoding: 'utf-8', timeout: 60000, stdio: 'pipe' })

  if (mdToHtml.status !== 0) {
    console.log(`  SKIP (md-to-html failed): ${mdToHtml.stderr || mdToHtml.stdout}`)
    results.push({ file: mdFile, status: 'SKIP', reason: 'md-to-html failed' })
    continue
  }

  // Step 2: html -> pptx
  const genPptx = spawnSync('node', [
    path.join(root, 'src', 'native-pptx', 'tools', 'gen-pptx.js'),
    htmlPath,
    pptxPath
  ], { encoding: 'utf-8', timeout: 180000, stdio: 'pipe' })

  if (genPptx.status !== 0) {
    console.log(`  SKIP (gen-pptx failed): ${(genPptx.stderr || genPptx.stdout).slice(0, 200)}`)
    results.push({ file: mdFile, status: 'SKIP', reason: 'gen-pptx failed' })
    continue
  }

  const pptxSize = fs.existsSync(pptxPath) ? fs.statSync(pptxPath).size : 0
  console.log(`  PPTX: ${pptxSize} bytes`)

  // Step 3: compare-visuals (output goes to dist/compare-<basename>/ automatically)
  const cmp = spawnSync('node', [
    path.join(root, 'src', 'native-pptx', 'tools', 'compare-visuals.js'),
    htmlPath,
    pptxPath
  ], { encoding: 'utf-8', timeout: 600000, stdio: 'pipe' })

  const output = cmp.stdout || ''
  const summaryMatch = output.match(/FAIL:\s*(\d+)\s+WARN:\s*(\d+)\s+OK:\s*(\d+)/)
  const failSlides = output.match(/FAILed slides:\s*(.+)/)?.[1] || ''
  
  if (summaryMatch) {
    const [, fail, warn, ok] = summaryMatch
    const total = parseInt(fail) + parseInt(warn) + parseInt(ok)
    const warnPct = total > 0 ? ((parseInt(warn) / total) * 100).toFixed(1) : '0'
    console.log(`  FAIL:${fail} WARN:${warn} OK:${ok} (WARN=${warnPct}%)${failSlides ? ' FAILed: ' + failSlides : ''}`)
    results.push({ file: mdFile, fail: parseInt(fail), warn: parseInt(warn), ok: parseInt(ok), total, warnPct, failSlides, exitCode: cmp.status })
  } else {
    console.log(`  No summary found. Exit: ${cmp.status}`)
    if (output) console.log(`  ${output.slice(-300)}`)
    results.push({ file: mdFile, status: 'ERROR', reason: 'no summary', exitCode: cmp.status })
  }
  console.log()
}

// Final summary
console.log('\n' + '='.repeat(60))
console.log('BATCH VR SUMMARY')
console.log('='.repeat(60))

const completed = results.filter(r => r.total)
const skipped = results.filter(r => r.status === 'SKIP' || r.status === 'ERROR')
const totalFail = completed.reduce((s, r) => s + r.fail, 0)
const totalWarn = completed.reduce((s, r) => s + r.warn, 0)
const totalOk = completed.reduce((s, r) => s + r.ok, 0)
const totalSlides = totalFail + totalWarn + totalOk

console.log(`Files: ${completed.length} completed, ${skipped.length} skipped`)
console.log(`Slides: ${totalSlides} total — FAIL:${totalFail} WARN:${totalWarn} OK:${totalOk}`)
if (totalSlides > 0) {
  console.log(`  FAIL rate: ${((totalFail / totalSlides) * 100).toFixed(1)}%`)
  console.log(`  WARN rate: ${((totalWarn / totalSlides) * 100).toFixed(1)}%`)
  console.log(`  OK   rate: ${((totalOk / totalSlides) * 100).toFixed(1)}%`)
}

console.log('\nPer-file results:')
for (const r of results) {
  if (r.total) {
    const mark = r.fail > 0 ? '✗' : r.warn > 0 ? '△' : '○'
    console.log(`  ${mark} ${r.file}: F${r.fail}/W${r.warn}/OK${r.ok} (${r.total} slides)${r.failSlides ? ' [FAIL: ' + r.failSlides + ']' : ''}`)
  } else {
    console.log(`  - ${r.file}: ${r.status} (${r.reason})`)
  }
}

// Write JSON report
const reportPath = path.join(distDir, 'batch-report.json')
fs.writeFileSync(reportPath, JSON.stringify(results, null, 2))
console.log(`\nReport: ${reportPath}`)

process.exit(totalFail > 0 ? 1 : 0)
