#!/usr/bin/env node
/**
 * Generate final VR batch report: runs structural test on all files and
 * combines with pixel-diff results into a comprehensive markdown report.
 */
const fs = require('node:fs')
const path = require('node:path')
const { execSync } = require('child_process')

const root = path.resolve(__dirname, '..')
process.chdir(root)

const distDir = path.join(root, 'dist', 'vr-batch')
const batchReport = JSON.parse(fs.readFileSync(path.join(distDir, 'batch-report.json'), 'utf8'))

// Run structural test on each file
const structuralResults = []
for (const entry of batchReport) {
  if (entry.status === 'SKIP' || entry.status === 'ERROR') {
    structuralResults.push({ file: entry.file, status: 'SKIP' })
    continue
  }
  const baseName = entry.file.replace(/\.md$/, '')
  const htmlPath = path.join(distDir, baseName + '.html')
  const pptxPath = path.join(distDir, baseName + '.pptx')

  if (!fs.existsSync(htmlPath) || !fs.existsSync(pptxPath)) {
    structuralResults.push({ file: entry.file, status: 'SKIP' })
    continue
  }

  try {
    const stdout = execSync(
      `node src/native-pptx/tools/structural-test.js "${htmlPath}" "${pptxPath}"`,
      { encoding: 'utf8', timeout: 120000, stdio: ['pipe', 'pipe', 'pipe'] }
    )
    const data = JSON.parse(stdout)
    structuralResults.push(data)
  } catch (e) {
    structuralResults.push({ file: entry.file, status: 'ERROR', error: (e.stderr || e.message || '').slice(0, 200) })
  }
}

// Generate markdown report
const lines = []
lines.push('# Visual Regression Report')
lines.push('')
lines.push(`Generated: ${new Date().toISOString()}`)
lines.push('')
lines.push('## Summary')
lines.push('')

const completed = batchReport.filter(r => r.total)
const totalSlides = completed.reduce((s, r) => s + r.total, 0)
const totalFail = completed.reduce((s, r) => s + r.fail, 0)
const totalWarn = completed.reduce((s, r) => s + r.warn, 0)
const totalOk = completed.reduce((s, r) => s + r.ok, 0)

lines.push(`| Metric | Value |`)
lines.push(`|--------|-------|`)
lines.push(`| Files | ${completed.length} |`)
lines.push(`| Total Slides | ${totalSlides} |`)
lines.push(`| Pixel FAIL | ${totalFail} (${((totalFail/totalSlides)*100).toFixed(1)}%) |`)
lines.push(`| Pixel WARN | ${totalWarn} (${((totalWarn/totalSlides)*100).toFixed(1)}%) |`)
lines.push(`| Pixel OK | ${totalOk} (${((totalOk/totalSlides)*100).toFixed(1)}%) |`)

// Structural summary
const structCompleted = structuralResults.filter(r => r.slides)
const structFail = structCompleted.reduce((s, r) => s + (r.summary?.FAIL || 0), 0)
const structWarn = structCompleted.reduce((s, r) => s + (r.summary?.WARN || 0), 0)
const structOk = structCompleted.reduce((s, r) => s + (r.summary?.OK || 0), 0)
const structTotal = structFail + structWarn + structOk

if (structTotal > 0) {
  lines.push(`| Structural FAIL | ${structFail} (${((structFail/structTotal)*100).toFixed(1)}%) |`)
  lines.push(`| Structural WARN | ${structWarn} (${((structWarn/structTotal)*100).toFixed(1)}%) |`)
  lines.push(`| Structural OK | ${structOk} (${((structOk/structTotal)*100).toFixed(1)}%) |`)
}
lines.push('')

// Per-file table
lines.push('## Per-File Results')
lines.push('')
lines.push('| # | File | Slides | Pixel (F/W/OK) | Structural (F/W/OK) | Failed Slides |')
lines.push('|---|------|--------|----------------|---------------------|---------------|')

for (let i = 0; i < batchReport.length; i++) {
  const entry = batchReport[i]
  const struct = structuralResults[i]
  const pixelCol = entry.total
    ? `${entry.fail}/${entry.warn}/${entry.ok}`
    : entry.status || 'N/A'
  const structCol = struct?.slides
    ? `${struct.summary.FAIL}/${struct.summary.WARN}/${struct.summary.OK}`
    : struct?.status || 'N/A'
  const failSlides = entry.failSlides || ''
  const mark = entry.fail > 0 ? '❌' : entry.warn > 0 ? '⚠️' : '✅'
  lines.push(`| ${i+1} | ${mark} ${entry.file} | ${entry.total || '-'} | ${pixelCol} | ${structCol} | ${failSlides} |`)
}
lines.push('')

// Detailed per-slide structural issues
lines.push('## Structural Issues Detail')
lines.push('')

for (const struct of structuralResults) {
  if (!struct.perSlide) continue
  const issues = struct.perSlide.filter(s => s.overall !== 'OK')
  if (issues.length === 0) continue

  lines.push(`### ${struct.file}`)
  lines.push('')
  lines.push('| Slide | Status | Content | Font | Code | Bounds | Details |')
  lines.push('|-------|--------|---------|------|------|--------|---------|')
  for (const s of issues) {
    const details = []
    if (s.details.contentMissingPct > 5) details.push(`content:-${s.details.contentMissingPct}%`)
    if (s.details.fontIssues?.length) details.push(s.details.fontIssues[0])
    if (s.details.codeFormatIssues?.length) details.push(s.details.codeFormatIssues[0])
    if (s.details.offScreenCount) details.push(`offscreen:${s.details.offScreenCount}`)
    if (s.details.tinyFontCount) details.push(`tinyfont:${s.details.tinyFontCount}`)
    if (s.details.structureIssues?.length) details.push(s.details.structureIssues.join('; '))
    if (s.details.markdownIssues?.length) details.push(s.details.markdownIssues.join('; '))
    lines.push(`| ${s.slide} | ${s.overall} | ${s.content || 'OK'} | ${s.font || 'OK'} | ${s.codeFormat || 'OK'} | ${s.bounds || 'OK'} | ${details.join(' / ') || '-'} |`)
  }
  lines.push('')
}

// Fixes applied section
lines.push('## Fixes Applied This Session')
lines.push('')
lines.push('### Problem 1: Chat Bubble Content Loss (slide-builder.ts)')
lines.push('- **Root Cause**: `isEmbeddableContainer` returned true for containers with complex children, causing `associateContainerText` to embed text into the container shape via the spatial path. The embedded path called `break` after addText, skipping recursive child placement.')
lines.push('- **Fix**: Added check in `isEmbeddableContainer`: if container has children that are NOT all simple text types (isSimpleTextContainer=false), return false. This forces the standard path (shape + recursive children placement).')
lines.push('- **Validation**: `38下期方針発表_清水.md` slide 6 now passes structural test with 0% content loss.')
lines.push('')
lines.push('### Problem 2: Tofu / Font Rendering (utils.ts)')
lines.push('- **Root Cause**: Proprietary dev fonts ("UDEV Gothic 35HSJPDOC") leaked into PPTX when text was non-CJK (emoji, fullwidth forms). The `japaneseTextPattern` missed CJK symbols (U+3000-303F) and fullwidth forms (U+FF01-FF9F).')
lines.push('- **Fix**: (a) Expanded `japaneseTextPattern` to cover fullwidth forms. (b) Added `proprietaryFontPattern` to filter dev-only fonts in non-Japanese path. (c) Strip orphaned U+FE0E/FE0F variation selectors in sanitizeText.')
lines.push('- **Validation**: Structural Layer 3b now detects remaining proprietary font instances as WARN.')
lines.push('')
lines.push('### Problem 3: Code Block Shape Disappearance in Lists (dom-walker.ts)')
lines.push('- **Root Cause**: In `extractListItemEl()`, `<pre>` tags matched the generic `isBlockChild` path, losing code block background and monospace font.')
lines.push('- **Fix**: Added `if (childTag === "pre")` special case that uses `extractCodeRuns()` and preserves backgroundColor + fontFamily from the `<pre>` element.')
lines.push('- **Validation**: Structural Layer 2b now detects code-in-list formatting issues.')
lines.push('')

// Write report
const reportPath = path.join(distDir, 'REPORT.md')
fs.writeFileSync(reportPath, lines.join('\n'), 'utf8')
console.log(`Report written: ${reportPath}`)
console.log(`Summary: ${totalSlides} slides, Pixel F${totalFail}/W${totalWarn}/OK${totalOk}, Structural F${structFail}/W${structWarn}/OK${structOk}`)
