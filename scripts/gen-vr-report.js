#!/usr/bin/env node
/**
 * Generate a single-page visual regression report HTML from batch test results.
 * Shows all slides side-by-side: Marp HTML (reference) | PPTX output.
 *
 * Usage:
 *   node scripts/gen-vr-report.js [output-path]
 * Default output: dist/vr-batch/vr-report.html
 */
const path = require('path')
const fs = require('fs')
const { pathToFileURL } = require('url')

const root = path.resolve(__dirname, '..')
const distDir = path.join(root, 'dist')
const batchJsonPath = path.join(distDir, 'vr-batch', 'batch-report.json')
const defaultOutputPath = process.argv[2] || path.join(distDir, 'vr-batch', 'vr-report.html')

if (!fs.existsSync(batchJsonPath)) {
  console.error('batch-report.json not found:', batchJsonPath)
  process.exit(1)
}

const batchReport = JSON.parse(fs.readFileSync(batchJsonPath, 'utf8'))

// Load structural test results if available
const structuralJsonPath = path.join(distDir, 'vr-batch', 'structural-report.json')
const structuralReport = fs.existsSync(structuralJsonPath)
  ? JSON.parse(fs.readFileSync(structuralJsonPath, 'utf8'))
  : []

/** Get structural test info for a file. Returns { FAIL, WARN, OK, failedSlides } */
function getStructuralInfo(basename) {
  const entry = structuralReport.find(e => e.file === basename)
  if (!entry) return null
  // Parse failedSlides like "3(C:31.0%)" into { slide, reasons }
  const parsed = (entry.failedSlides || []).map(s => {
    const m = s.match(/^(\d+)\((.+)\)$/)
    return m ? { slide: parseInt(m[1]), reasons: m[2] } : null
  }).filter(Boolean)
  return { ...entry, parsedFails: parsed }
}

/** Convert absolute Windows path to file:// URL for use in HTML src attributes. */
function imgUrl(absPath) {
  return pathToFileURL(absPath).href
}

/**
 * Parse per-slide status from compare-report.html.
 * Returns a map: slideNum (1-based) -> 'FAIL' | 'WARN' | 'OK' | 'MISSING'
 */
function parseSlideStatus(compareDir) {
  const reportPath = path.join(compareDir, 'compare-report.html')
  if (!fs.existsSync(reportPath)) return {}
  const html = fs.readFileSync(reportPath, 'utf8')
  const statusMap = {}
  // Match pattern: <number>\n<br><small>STATUS
  const re = /(\d+)<br><small>(FAIL|WARN|OK|MISSING)/g
  let m
  while ((m = re.exec(html)) !== null) {
    statusMap[parseInt(m[1])] = m[2]
  }
  return statusMap
}

/**
 * Get list of slide image paths for a given compare directory.
 * Returns sorted array of { num, htmlImg, pptxImg } objects.
 */
function getSlideImages(compareDir) {
  if (!fs.existsSync(compareDir)) return []
  const files = fs.readdirSync(compareDir)
  const htmlSlides = files.filter(f => /^html-slide-\d+\.png$/.test(f)).sort()
  return htmlSlides.map(f => {
    const numStr = f.match(/(\d{3})\.png$/)?.[1] ?? '000'
    const num = parseInt(numStr)
    return {
      num,
      htmlImg: path.join(compareDir, f),
      pptxImg: path.join(compareDir, `pptx-slide-${numStr}.png`),
    }
  })
}

const CSS = `
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #181818; color: #e0e0e0; }

/* ── TOC (sticky header) ───────────────────────────────────────── */
#toc {
  position: sticky; top: 0; z-index: 100;
  background: #232323; border-bottom: 2px solid #3a3a3a;
  padding: 8px 14px;
  display: flex; flex-wrap: wrap; align-items: center; gap: 5px;
}
#toc .toc-label { font-size: 11px; font-weight: bold; color: #888; margin-right: 4px; }
#toc a {
  font-size: 10px; padding: 2px 7px; border-radius: 3px;
  border: 1px solid #555; color: #bbb; text-decoration: none;
  white-space: nowrap;
}
#toc a:hover { background: #3a3a3a; }
#toc a.has-fail  { border-color: #f55; color: #f99; }
#toc a.has-warn  { border-color: #fb0; color: #fd9; }
#toc a.all-ok    { border-color: #5c5; color: #9d9; }

/* ── Summary bar ────────────────────────────────────────────────── */
#summary {
  background: #1e1e1e; border-bottom: 1px solid #333;
  padding: 10px 16px; font-size: 12px; color: #aaa;
  display: flex; gap: 20px; flex-wrap: wrap;
}
#summary strong { color: #ddd; }

/* ── File sections ─────────────────────────────────────────────── */
.file-section { padding: 20px 16px 6px; border-top: 3px solid #333; }
.file-header {
  display: flex; align-items: center; gap: 10px; margin-bottom: 10px; flex-wrap: wrap;
}
.file-title { font-size: 15px; font-weight: bold; color: #ddd; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.badge { display: inline-block; padding: 2px 9px; border-radius: 3px; font-weight: bold; font-size: 11px; }
.bg-fail   { background: #c33; color: #fff; }
.bg-warn   { background: #b80; color: #fff; }
.bg-ok     { background: #2a7a2a; color: #fff; }
.small-stat { font-size: 11px; color: #777; }

/* ── Slide grid ─────────────────────────────────────────────────── */
.slides-grid { display: flex; flex-direction: column; gap: 8px; }
.slide-row {
  display: flex; gap: 0; background: #252525;
  border-radius: 5px; overflow: hidden;
  border: 1px solid transparent;
}
.slide-row.status-fail { border-color: #833; }
.slide-row.status-warn { border-color: #773; }

.slide-num {
  width: 52px; flex-shrink: 0;
  display: flex; flex-direction: column; align-items: center; justify-content: center;
  font-size: 14px; font-weight: bold; padding: 6px 2px; gap: 3px;
}
.slide-num .status-label { font-size: 9px; text-transform: uppercase; }
.slide-num.s-fail   { background: #c33; color: #fff; }
.slide-num.s-warn   { background: #b80; color: #fff; }
.slide-num.s-ok     { background: #2a7a2a; color: #fff; }
.slide-num.s-missing{ background: #666; color: #fff; }

.slide-imgs { display: flex; flex: 1; gap: 0; min-width: 0; }
.img-col { flex: 1; padding: 3px; display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.img-col img { width: 100%; display: block; border: 1px solid #444; border-radius: 2px; }
.img-col .img-label { font-size: 9px; color: #666; text-align: center; padding-bottom: 1px; }
.img-sep { width: 2px; background: #333; flex-shrink: 0; }

/* ── Structural test annotations ──────────────────────────────────── */
.struct-badge { display: inline-block; padding: 1px 5px; border-radius: 2px; font-size: 9px; font-weight: bold; margin-left: 4px; }
.struct-badge.sb-fail { background: #a22; color: #fff; }
.struct-badge.sb-warn { background: #885500; color: #fff; }
.struct-info { font-size: 10px; color: #aaa; margin-top: 2px; }
.struct-reasons { font-size: 9px; color: #f99; margin-top: 1px; }
.layer-tag { display: inline-block; padding: 0 4px; border-radius: 2px; font-size: 8px; font-weight: bold; margin-right: 2px; }
.layer-tag.lt-c { background: #633; color: #faa; }
.layer-tag.lt-s { background: #653; color: #fda; }
.layer-tag.lt-t { background: #636; color: #faf; }
.layer-tag.lt-b { background: #366; color: #aff; }
`

const generatedAt = new Date().toLocaleString('ja-JP')

const parts = []
parts.push(`<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>VR Report — ${generatedAt}</title>
<style>${CSS}</style>
</head>
<body>
`)

// ── TOC ─────────────────────────────────────────────────────────────
parts.push(`<div id="toc">\n<span class="toc-label">JUMP:</span>\n`)
let totalFiles = 0, totalSlides = 0, totalFail = 0, totalWarn = 0, totalOk = 0

for (const entry of batchReport) {
  if (!entry.total) continue
  totalFiles++
  totalSlides += entry.total
  totalFail += entry.fail
  totalWarn += entry.warn
  totalOk += entry.ok
  const basename = entry.file.replace(/\.md$/, '')
  const cls = entry.fail > 0 ? 'has-fail' : entry.warn > 0 ? 'has-warn' : 'all-ok'
  const shortName = basename.length > 28 ? basename.slice(0, 28) + '…' : basename
  parts.push(`<a href="#file-${encodeURIComponent(basename)}" class="${cls}" title="${basename}\nF${entry.fail} W${entry.warn} OK${entry.ok} / ${entry.total}slides">${shortName} (F${entry.fail})</a>\n`)
}
parts.push(`</div>\n`)

// ── Summary bar ─────────────────────────────────────────────────────
const failPct = totalSlides > 0 ? ((totalFail / totalSlides) * 100).toFixed(1) : '0'
const warnPct = totalSlides > 0 ? ((totalWarn / totalSlides) * 100).toFixed(1) : '0'
const okPct   = totalSlides > 0 ? ((totalOk   / totalSlides) * 100).toFixed(1) : '0'

// Structural test totals
const sFail = structuralReport.reduce((s, e) => s + (e.FAIL || 0), 0)
const sWarn = structuralReport.reduce((s, e) => s + (e.WARN || 0), 0)
const sOk   = structuralReport.reduce((s, e) => s + (e.OK || 0), 0)
const sTotal = sFail + sWarn + sOk

parts.push(`<div id="summary">
  <span><strong>Generated:</strong> ${generatedAt}</span>
  <span><strong>Files:</strong> ${totalFiles}</span>
  <span><strong>Slides:</strong> ${totalSlides}</span>
  <span style="color:#f99"><strong>FAIL:</strong> ${totalFail} (${failPct}%)</span>
  <span style="color:#fd9"><strong>WARN:</strong> ${totalWarn} (${warnPct}%)</span>
  <span style="color:#9d9"><strong>OK:</strong>   ${totalOk} (${okPct}%)</span>
  ${sTotal > 0 ? `<span style="border-left:1px solid #555;padding-left:12px"><strong>Structural:</strong> <span style="color:#f99">F${sFail}</span> / <span style="color:#fd9">W${sWarn}</span> / <span style="color:#9d9">OK${sOk}</span></span>` : ''}
</div>\n`)

// ── File sections ────────────────────────────────────────────────────
for (const entry of batchReport) {
  if (!entry.total) continue
  const basename = entry.file.replace(/\.md$/, '')
  const compareDir = path.join(distDir, `compare-${basename}`)
  const slideStatus = parseSlideStatus(compareDir)
  const slides = getSlideImages(compareDir)
  const structInfo = getStructuralInfo(basename)

  const overallBadgeCls = entry.fail > 0 ? 'bg-fail' : entry.warn > 0 ? 'bg-warn' : 'bg-ok'
  const overallLabel    = entry.fail > 0 ? '✗ FAIL' : entry.warn > 0 ? '△ WARN' : '○ OK'
  const id = encodeURIComponent(basename)

  // Structural test badge
  let structBadge = ''
  if (structInfo) {
    const sCls = structInfo.FAIL > 0 ? 'sb-fail' : structInfo.WARN > 0 ? 'sb-warn' : ''
    const sLabel = structInfo.FAIL > 0 ? `Struct:F${structInfo.FAIL}` : `Struct:W${structInfo.WARN}`
    if (sCls) structBadge = `<span class="struct-badge ${sCls}">${sLabel}</span>`
  }

  parts.push(`<div class="file-section" id="file-${id}">
<div class="file-header">
  <span class="file-title" title="${basename}">${basename}</span>
  <span class="badge ${overallBadgeCls}">${overallLabel}</span>
  ${structBadge}
  <span class="small-stat">Pixel: F${entry.fail}/W${entry.warn}/OK${entry.ok}${structInfo ? ` | Structural: F${structInfo.FAIL}/W${structInfo.WARN}/OK${structInfo.OK}` : ''} — ${entry.total} slides</span>
</div>
<div class="slides-grid">
`)

  for (const slide of slides) {
    const status = slideStatus[slide.num] ?? 'OK'
    const numCls  = `s-${status.toLowerCase()}`
    const rowCls  = status === 'FAIL' ? ' status-fail' : status === 'WARN' ? ' status-warn' : ''
    const htmlSrc = fs.existsSync(slide.htmlImg) ? imgUrl(slide.htmlImg) : ''
    const pptxSrc = fs.existsSync(slide.pptxImg) ? imgUrl(slide.pptxImg) : ''

    // Check if this slide has structural test failures
    let structAnnotation = ''
    if (structInfo) {
      const failEntry = structInfo.parsedFails.find(f => f.slide === slide.num)
      if (failEntry) {
        const reasons = failEntry.reasons
        const tags = []
        if (reasons.includes('C:')) tags.push('<span class="layer-tag lt-c">C</span>')
        if (reasons.includes('S:')) tags.push('<span class="layer-tag lt-s">S</span>')
        if (reasons.includes('T'))  tags.push('<span class="layer-tag lt-t">T</span>')
        if (reasons.includes('B'))  tags.push('<span class="layer-tag lt-b">B</span>')
        structAnnotation = `<div class="struct-reasons">${tags.join('')} ${reasons}</div>`
      }
    }

    parts.push(`<div class="slide-row${rowCls}">
  <div class="slide-num ${numCls}">${slide.num}<span class="status-label">${status}</span>${structAnnotation}</div>
  <div class="slide-imgs">
    <div class="img-col">
      ${htmlSrc ? `<img src="${htmlSrc}" loading="lazy" alt="HTML ${slide.num}">` : '<div style="color:#666;text-align:center;padding:16px">—</div>'}
      <div class="img-label">Marp HTML (ref)</div>
    </div>
    <div class="img-sep"></div>
    <div class="img-col">
      ${pptxSrc ? `<img src="${pptxSrc}" loading="lazy" alt="PPTX ${slide.num}">` : '<div style="color:#666;text-align:center;padding:16px">—</div>'}
      <div class="img-label">PPTX output</div>
    </div>
  </div>
</div>
`)
  }

  parts.push(`</div><!-- slides-grid -->
</div><!-- file-section -->
`)
}

parts.push(`</body>
</html>
`)

const html = parts.join('')
fs.mkdirSync(path.dirname(defaultOutputPath), { recursive: true })
fs.writeFileSync(defaultOutputPath, html, 'utf8')
console.log(`Report written: ${defaultOutputPath}`)
console.log(`  Size: ${(html.length / 1024).toFixed(1)} KB`)
console.log(`  Files: ${totalFiles}, Slides: ${totalSlides}`)
console.log(`  FAIL: ${totalFail} (${failPct}%), WARN: ${totalWarn} (${warnPct}%), OK: ${totalOk} (${okPct}%)`)
