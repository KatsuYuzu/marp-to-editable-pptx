#!/usr/bin/env node
/**
 * Generates a single all-in-one HTML comparison report aggregating every
 * compare-<name>/ directory under dist/.
 *
 * Usage:
 *   node scripts/gen-all-compare-report.js [dist-dir]
 *
 * Output:
 *   dist/all-compare-report.html
 *
 * Image paths are relative so the HTML can be opened directly from the
 * dist/ folder without a web server.
 *
 * Data sources:
 *   dist/vr-batch/batch-report.json     — per-file pixel FAIL/WARN/OK + failSlides
 *   dist/vr-batch/structural-report.json — per-file structural issues + failedSlides
 *   dist/compare-<name>/html-slide-NNN.png
 *   dist/compare-<name>/pptx-slide-NNN.png
 *   dist/compare-<name>/diff-slide-NNN.png
 */
const fs = require('node:fs')
const path = require('node:path')

const root = path.resolve(__dirname, '..')
const distDir = path.resolve(process.argv[2] || path.join(root, 'dist'))
const batchDir = path.join(distDir, 'vr-batch')

// ── Load batch pixel report ────────────────────────────────────────────────
const batchReportPath = path.join(batchDir, 'batch-report.json')
/** @type {Array<{file:string,fail?:number,warn?:number,ok?:number,total?:number,failSlides?:string,status?:string}>} */
let batchReport = []
try {
  batchReport = JSON.parse(fs.readFileSync(batchReportPath, 'utf8'))
} catch {
  console.warn('batch-report.json not found — pixel status will be unknown')
}

// ── Load structural report ─────────────────────────────────────────────────
const structReportPath = path.join(batchDir, 'structural-report.json')
/**
 * @type {Array<{file:string,FAIL:number,WARN:number,OK:number,slides:number,failedSlides:string[]}>}
 */
let structReport = []
try {
  structReport = JSON.parse(fs.readFileSync(structReportPath, 'utf8'))
} catch {
  console.warn('structural-report.json not found — structural status will be unknown')
}

// ── Discover compare directories ─────────────────────────────────────────
const compareDirs = fs
  .readdirSync(distDir)
  .filter(n => n.startsWith('compare-'))
  .filter(n => {
    const full = path.join(distDir, n)
    return fs.statSync(full).isDirectory()
  })
  .sort()

if (compareDirs.length === 0) {
  console.error('No compare-* directories found in', distDir)
  process.exit(1)
}
console.log(`Found ${compareDirs.length} compare dirs`)

// ── Build per-file data ───────────────────────────────────────────────────
/**
 * Parse a pixel fail-slides string like "1, 3, 7" into a Set<number>.
 * @param {string|undefined} s
 * @returns {Set<number>}
 */
function parseFailSlides(s) {
  const result = new Set()
  if (!s) return result
  for (const part of s.split(',')) {
    const n = parseInt(part.trim(), 10)
    if (!isNaN(n)) result.add(n)
  }
  return result
}

/**
 * Parse structural failedSlides strings like "3(C:31.0%+S:list: HTML=1, PPTX=0+T)".
 * Format: <slideNum>(<reason1>+<reason2>+...)
 * Reason codes: C:<pct>%  S:<issues>  T  B  M  V  F  CF
 * Returns a Map<slideNum, reasonCode[]>.
 * @param {string[]|undefined} arr
 * @returns {Map<number, string[]>}
 */
function parseStructFails(arr) {
  const map = new Map()
  if (!arr) return map
  for (const entry of arr) {
    // Format: "1(C:45.2%+T)" or "5(S:list: HTML=1, PPTX=0+M)"
    const parenIdx = entry.indexOf('(')
    if (parenIdx === -1) {
      // plain number
      const n = parseInt(entry, 10)
      if (!isNaN(n)) {
        if (!map.has(n)) map.set(n, [])
        map.get(n).push('FAIL')
      }
      continue
    }
    const n = parseInt(entry.slice(0, parenIdx), 10)
    if (isNaN(n)) continue
    const reasons = entry.slice(parenIdx + 1, entry.endsWith(')') ? entry.length - 1 : entry.length)
    // Split on '+' but preserve 'S:<issues>' multi-word strings
    // Approach: scan for known prefixes and split properly
    const parts = []
    let cur = ''
    let i = 0
    while (i < reasons.length) {
      if (reasons[i] === '+' && cur.length > 0) {
        // Check if next token is a known prefix
        const rest = reasons.slice(i + 1)
        const isNewToken = /^(C:|S:|T\b|B\b|M\b|V\b|F\b|CF\b)/.test(rest)
        if (isNewToken) {
          parts.push(cur.trim())
          cur = ''
          i++
          continue
        }
      }
      cur += reasons[i]
      i++
    }
    if (cur.trim()) parts.push(cur.trim())
    if (!map.has(n)) map.set(n, [])
    map.get(n).push(...parts)
  }
  return map
}

/**
 * Decode human-readable label for a structural failure reason code.
 * @param {string} code
 * @returns {{label: string, isCritical: boolean}}
 */
function describeStructCode(code) {
  if (code === 'T') return { label: 'Tofu (missing glyph)', isCritical: false }
  if (code === 'B') return { label: 'Shape out of bounds', isCritical: false }
  if (code === 'M') return { label: '⚠ Raw markdown leaked into PPTX', isCritical: true }
  if (code === 'V') return { label: 'Visual coverage loss', isCritical: true }
  if (code === 'F') return { label: 'Proprietary/unavailable font', isCritical: false }
  if (code === 'CF') return { label: 'Code block formatting lost', isCritical: false }
  if (code === 'FAIL') return { label: 'Structural FAIL', isCritical: false }
  if (code.startsWith('C:')) {
    const pct = parseFloat(code.slice(2))
    const isCritical = pct > 50
    return { label: `Content loss ${pct.toFixed(1)}%`, isCritical }
  }
  if (code.startsWith('S:')) {
    const detail = code.slice(2)
    const isCritical = detail.includes('list: HTML') && detail.includes('PPTX=0')
    return { label: `Structure: ${detail}`, isCritical }
  }
  return { label: code, isCritical: false }
}

const esc = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

// ── Generate report ───────────────────────────────────────────────────────

const fileSections = []
let grandTotal = 0
let grandFail = 0
let grandWarn = 0
let grandOk = 0
let grandCritical = 0 // human-critical: content loss > 50% or raw markdown

const tocEntries = []

for (const dir of compareDirs) {
  const name = dir.replace(/^compare-/, '')
  const dirPath = path.join(distDir, dir)

  // Collect slide image sets
  const allFiles = fs.readdirSync(dirPath)
  const htmlPngs = allFiles.filter(f => /^html-slide-\d+\.png$/.test(f)).sort()
  const pptxPngs = allFiles.filter(f => /^pptx-slide-\d+\.png$/.test(f)).sort()
  const diffPngs = new Set(allFiles.filter(f => /^diff-slide-\d+\.png$/.test(f)))

  const maxSlides = Math.max(htmlPngs.length, pptxPngs.length)
  if (maxSlides === 0) continue

  // Look up pixel data
  const batchEntry = batchReport.find(r => r.file === name + '.md' || r.file === name)
  const pixelFail = parseFailSlides(batchEntry?.failSlides)
  const pixelFails = batchEntry?.fail ?? 0
  const pixelWarns = batchEntry?.warn ?? 0
  const pixelOks = batchEntry?.ok ?? 0

  // Look up structural data
  const structEntry = structReport.find(r => r.file === name)
  const structFails = parseStructFails(structEntry?.failedSlides)
  const structFailCount = structEntry?.FAIL ?? 0
  const structWarnCount = structEntry?.WARN ?? 0

  // Detect human-critical issues across slides
  let fileCritical = false
  const criticalSlides = new Set()

  for (const [slideNum, codes] of structFails.entries()) {
    for (const code of codes) {
      const { isCritical } = describeStructCode(code)
      if (isCritical) {
        criticalSlides.add(slideNum)
        fileCritical = true
        grandCritical++
      }
    }
  }

  grandTotal += maxSlides
  grandFail += pixelFails
  grandWarn += pixelWarns
  grandOk += pixelOks

  const anchId = `file-${encodeURIComponent(name)}`

  const slideRows = []
  for (let i = 0; i < maxSlides; i++) {
    const n = i + 1
    const pad = String(n).padStart(3, '0')
    const htmlImg = htmlPngs[i] ? `${dir}/html-slide-${pad}.png` : null
    const pptxImg = pptxPngs[i] ? `${dir}/pptx-slide-${pad}.png` : null
    const diffImg = diffPngs.has(`diff-slide-${pad}.png`) ? `${dir}/diff-slide-${pad}.png` : null

    const isPixelFail = pixelFail.has(n)
    const structCodes = structFails.get(n) || []
    const hasCritical = criticalSlides.has(n)

    // Border color: red for pixel FAIL, orange for struct FAIL, default otherwise
    const borderColor = isPixelFail ? '#e33' : structCodes.length > 0 ? '#e90' : '#ccc'
    const rowBg = isPixelFail ? '#fff5f5' : structCodes.length > 0 ? '#fffbf0' : ''

    // Build badge HTML for structural issues
    let structBadges = ''
    for (const code of structCodes) {
      const { label, isCritical } = describeStructCode(code)
      const bgColor = isCritical ? '#c00' : '#e90'
      structBadges += `<span class="badge" style="background:${bgColor}" title="${esc(label)}">⚠ ${esc(label)}</span> `
    }

    slideRows.push(`
      <tr id="${anchId}-slide-${n}" style="background:${rowBg}">
        <td class="slide-num" style="border-left:4px solid ${borderColor}">
          <strong>${n}</strong>
          ${isPixelFail ? '<br><span class="badge fail">PIXEL FAIL</span>' : ''}
          ${structCodes.length > 0 ? '<br><span class="badge struct">STRUCT FAIL</span>' : ''}
          ${hasCritical ? '<br><span class="badge critical">⚠ CRITICAL</span>' : ''}
        </td>
        <td class="img-cell">
          ${htmlImg
            ? `<img src="${esc(htmlImg)}" alt="HTML slide ${n}" loading="lazy">`
            : '<em style="color:#aaa">missing</em>'}
          <div class="img-label">Marp HTML</div>
        </td>
        <td class="img-cell">
          ${pptxImg
            ? `<img src="${esc(pptxImg)}" alt="PPTX slide ${n}" loading="lazy">`
            : '<em style="color:#aaa">missing</em>'}
          <div class="img-label">PPTX (native)</div>
          ${structBadges ? `<div class="struct-badges">${structBadges}</div>` : ''}
        </td>
        <td class="img-cell">
          ${diffImg
            ? `<img src="${esc(diffImg)}" alt="diff ${n}" loading="lazy">`
            : '<em style="color:#aaa">—</em>'}
          <div class="img-label">Pixel diff</div>
        </td>
      </tr>`)
  }

  const hasIssues = pixelFails > 0 || structFailCount > 0
  const sectionBg = fileCritical ? '#fff8e1' : ''
  const sectionBorderColor = pixelFails > 0 || structFailCount > 0 ? '#e33' : '#4a9'

  tocEntries.push(`<li><a href="#${anchId}">${esc(name)}</a>
    <span class="toc-stats">
      ${pixelFails > 0 ? `<span class="badge fail">F${pixelFails}</span>` : ''}
      ${pixelWarns > 0 ? `<span class="badge warn">W${pixelWarns}</span>` : ''}
      ${pixelOks > 0 ? `<span class="badge ok">OK${pixelOks}</span>` : ''}
      ${fileCritical ? '<span class="badge critical">⚠ CRITICAL</span>' : ''}
    </span>
  </li>`)

  fileSections.push(`
  <section id="${anchId}" style="background:${sectionBg}; border-left:6px solid ${sectionBorderColor}; padding:0 0 32px 12px; margin-bottom:32px">
    <h2 style="margin:0 0 8px 0; font-size:16px">${esc(name)}</h2>
    <div class="file-summary">
      Slides: ${maxSlides}
      &nbsp;|&nbsp;Pixel: <span class="badge fail">FAIL ${pixelFails}</span>
      <span class="badge warn">WARN ${pixelWarns}</span>
      <span class="badge ok">OK ${pixelOks}</span>
      &nbsp;|&nbsp;Structural: <span class="badge ${structFailCount > 0 ? 'fail' : 'ok'}">FAIL ${structFailCount}</span>
      <span class="badge ${structWarnCount > 0 ? 'warn' : 'ok'}">WARN ${structWarnCount}</span>
      ${fileCritical ? '&nbsp;|&nbsp;<span class="badge critical">⚠ Human-critical issues detected</span>' : ''}
      &nbsp;|&nbsp;<a href="${esc(dir + '/compare-report.html')}" target="_blank">Per-file report ↗</a>
    </div>
    <table class="slides-table">
      <thead>
        <tr>
          <th style="width:90px"># / Status</th>
          <th>Source (HTML)</th>
          <th>Output (PPTX)</th>
          <th>Pixel diff</th>
        </tr>
      </thead>
      <tbody>
        ${slideRows.join('\n')}
      </tbody>
    </table>
  </section>`)
}

// ── Stitch HTML ───────────────────────────────────────────────────────────
const genDate = new Date().toLocaleString('ja-JP')
const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>All Slides Comparison Report — ${genDate}</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: 'Segoe UI', sans-serif;
    margin: 0;
    background: #f2f4f6;
    color: #222;
  }
  header {
    background: #1a2030;
    color: #fff;
    padding: 16px 24px;
    position: sticky;
    top: 0;
    z-index: 100;
    display: flex;
    align-items: center;
    gap: 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,.4);
  }
  header h1 { margin: 0; font-size: 15px; font-weight: 600; }
  .grand-stats { font-size: 13px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
  .main { display: flex; }
  nav {
    width: 260px;
    min-width: 200px;
    background: #fff;
    border-right: 1px solid #ddd;
    padding: 16px 8px;
    font-size: 12px;
    height: 100vh;
    overflow-y: auto;
    position: sticky;
    top: 52px;
    align-self: flex-start;
    flex-shrink: 0;
  }
  nav h2 { font-size: 12px; color: #888; text-transform: uppercase; margin: 0 0 6px 0; }
  nav ol { margin: 0; padding-left: 18px; }
  nav li { margin-bottom: 6px; line-height: 1.3; }
  nav a { color: #346; text-decoration: none; word-break: break-all; }
  nav a:hover { text-decoration: underline; }
  .toc-stats { display: flex; flex-wrap: wrap; gap: 2px; margin-top: 2px; }
  article {
    flex: 1;
    min-width: 0;
    padding: 16px 20px 40px 20px;
  }
  .badge {
    display: inline-block;
    padding: 1px 6px;
    border-radius: 3px;
    color: #fff;
    font-size: 10px;
    font-weight: bold;
    white-space: nowrap;
  }
  .fail     { background: #d33; }
  .warn     { background: #e90; }
  .ok       { background: #3a8; }
  .struct   { background: #e50; }
  .critical { background: #800; }
  .file-summary {
    font-size: 13px;
    margin-bottom: 8px;
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    align-items: center;
  }
  .slides-table {
    border-collapse: collapse;
    table-layout: fixed;
    width: 100%;
  }
  .slides-table th {
    background: #2a3548;
    color: #dde;
    padding: 6px 8px;
    text-align: left;
    font-size: 12px;
    position: sticky;
    top: 52px;
    z-index: 5;
  }
  .slides-table td { vertical-align: top; padding: 4px; }
  .slide-num {
    width: 90px;
    font-size: 13px;
    text-align: center;
    vertical-align: middle;
    padding: 8px 4px;
  }
  .img-cell { width: calc((100% - 90px) / 3); }
  .img-cell img {
    width: 100%;
    height: auto;
    display: block;
    border: 1px solid #ccc;
    border-radius: 2px;
  }
  .img-label { font-size: 10px; color: #666; margin-top: 2px; text-align: center; }
  .struct-badges { margin-top: 4px; display: flex; flex-wrap: wrap; gap: 2px; }
  section h2 { padding: 8px 0 0 0; }
  .summary-box {
    background: #fff;
    border: 1px solid #ddd;
    border-radius: 6px;
    padding: 12px 16px;
    margin-bottom: 24px;
    font-size: 14px;
  }
  .summary-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 16px;
    margin-top: 10px;
  }
  .summary-card {
    background: #f8f9fb;
    border: 1px solid #e0e4ea;
    border-radius: 4px;
    padding: 8px 16px;
    min-width: 120px;
    text-align: center;
  }
  .summary-card .num { font-size: 28px; font-weight: bold; line-height: 1; }
  .summary-card .lbl { font-size: 11px; color: #666; margin-top: 4px; }
  .critical-box {
    background: #fff8e1;
    border: 2px solid #f9a825;
    border-radius: 6px;
    padding: 10px 14px;
    margin-bottom: 16px;
    font-size: 13px;
  }
  .critical-box strong { color: #c62828; }
</style>
</head>
<body>
<header>
  <h1>All Slides Comparison Report</h1>
  <div class="grand-stats">
    <span>${compareDirs.length} files &middot; ${grandTotal} slides</span>
    <span class="badge fail">PIXEL FAIL ${grandFail}</span>
    <span class="badge warn">PIXEL WARN ${grandWarn}</span>
    <span class="badge ok">PIXEL OK ${grandOk}</span>
    ${grandCritical > 0 ? `<span class="badge critical">⚠ ${grandCritical} CRITICAL</span>` : ''}
    <span style="color:#aaa;font-size:11px">Generated: ${genDate}</span>
  </div>
</header>

<div class="main">
  <nav>
    <h2>Files (${compareDirs.length})</h2>
    <ol>
      ${tocEntries.join('\n')}
    </ol>
  </nav>
  <article>
    <div class="summary-box">
      <strong>Grand summary</strong>
      <div class="summary-grid">
        <div class="summary-card">
          <div class="num">${compareDirs.length}</div>
          <div class="lbl">Files</div>
        </div>
        <div class="summary-card">
          <div class="num">${grandTotal}</div>
          <div class="lbl">Total slides</div>
        </div>
        <div class="summary-card" style="border-color:#d33">
          <div class="num" style="color:#d33">${grandFail}</div>
          <div class="lbl">Pixel FAIL</div>
        </div>
        <div class="summary-card" style="border-color:#e90">
          <div class="num" style="color:#e90">${grandWarn}</div>
          <div class="lbl">Pixel WARN</div>
        </div>
        <div class="summary-card" style="border-color:#3a8">
          <div class="num" style="color:#3a8">${grandOk}</div>
          <div class="lbl">Pixel OK</div>
        </div>
        ${grandCritical > 0 ? `
        <div class="summary-card" style="border-color:#800">
          <div class="num" style="color:#800">${grandCritical}</div>
          <div class="lbl">Critical (human-visible)</div>
        </div>` : ''}
      </div>
    </div>

    ${grandCritical > 0 ? `
    <div class="critical-box">
      <strong>⚠ Human-critical patterns detected</strong>
      — These slides may look acceptable to pixel-diff but have catastrophic
      content differences that a human reviewer would immediately notice:
      content loss &gt;50%, raw markdown syntax rendering as literal text,
      or list bullet markers completely missing.
      Slides flagged with <span class="badge critical">⚠ CRITICAL</span> require manual inspection.
    </div>` : ''}

    ${fileSections.join('\n')}
  </article>
</div>
</body>
</html>`

const outPath = path.join(distDir, 'all-compare-report.html')
fs.writeFileSync(outPath, html, 'utf8')
console.log(`Report written: ${outPath}`)
console.log(`  ${compareDirs.length} files, ${grandTotal} slides`)
console.log(`  PIXEL FAIL:${grandFail} WARN:${grandWarn} OK:${grandOk}`)
if (grandCritical > 0) console.log(`  ⚠ ${grandCritical} human-critical issues`)
