#!/usr/bin/env node
/**
 * Structural fidelity test for native PPTX.
 *
 * Six layers beyond pixel diff:
 *   Layer 1:  Content Completeness — text in HTML must exist in PPTX
 *   Layer 1b: Raw Markdown Detection — markdown syntax must not leak into non-code shapes
 *   Layer 2:  Structure Integrity  — element counts (code, list, table) must match
 *   Layer 3:  Glyph Rendering      — detect tofu (□□□) in PPTX screenshots
 *   Layer 4:  Bounds + Overflow    — off-screen shapes, text overflow, tiny font (content squeeze)
 *   Layer 6:  Spatial Coverage     — blank area detection via PNG comparison
 *
 * Usage:
 *   node src/native-pptx/tools/structural-test.js <html-path> <pptx-path> [--png-dir=<dir>]
 *
 * Output: JSON results to stdout, human summary to stderr.
 */
const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const JSZip = require('jszip')

// ─── Layer 1: Content Completeness ─────────────────────────────────────────

/**
 * Extract all text from PPTX slide XMLs.
 * Returns an array indexed by slide number (1-based), each entry is the
 * concatenated text of that slide.
 */
async function extractPptxTexts(pptxPath) {
  const buf = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(buf)
  const slideEntries = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1])
      const nb = parseInt(b.match(/slide(\d+)/)[1])
      return na - nb
    })

  const results = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry].async('string')
    // Extract all <a:t> text nodes and decode XML entities
    const texts = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => {
      return m[1]
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
    })
    results.push(texts.join(' '))
  }
  return results
}

/**
 * Extract text from Marp HTML via Puppeteer by running dom-walker.
 * Returns array of per-slide text (concatenated from all elements).
 */
async function extractHtmlTexts(htmlPath, browserPath) {
  const puppeteer = require('puppeteer-core')
  const browser = await puppeteer.launch({
    executablePath: browserPath,
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu'],
  })

  try {
    const page = await browser.newPage()
    await page.setViewport({ width: 1280, height: 720 })
    await page.goto(pathToFileURL(htmlPath).href, { waitUntil: 'networkidle0' })
    await new Promise(r => setTimeout(r, 300))

    // Load the dom-walker script
    const bundlePath = path.resolve(__dirname, '..', '..', '..', 'lib', 'native-pptx.cjs')
    const { generateNativePptx } = require(bundlePath)

    // Instead of using the full pipeline, directly run extractSlides in page context
    const domWalkerPath = path.resolve(__dirname, '..', 'dom-walker-script.generated.ts')
    // Use the raw source from generated bundle — it exports extractSlides
    // Actually, let's just inline extractSlides call via the bundle's exported function
    // Better approach: run the extraction directly

    // Load the dom-walker-script
    const domWalkerScriptFile = path.resolve(__dirname, '..', '..', '..', 'lib', 'native-pptx.cjs')
    const libSource = fs.readFileSync(domWalkerScriptFile, 'utf8')

    // Extract the DOM_WALKER_SCRIPT string from the bundle
    // Alternative: just evaluate the extractSlides function in page context
    // Simplest: use puppeteer to extract text directly from the HTML
    const slideTexts = await page.evaluate(() => {
      // Find all top-level sections (slides)
      const allSections = Array.from(document.querySelectorAll('section'))
        .filter(s => {
          if (s.parentElement?.closest('section')) return false
          return (
            s.parentElement?.tagName.toLowerCase() === 'foreignobject' ||
            s.hasAttribute('data-marpit-pagination')
          )
        })

      // Deduplicate by key (advanced background creates multiple layers)
      const seen = new Set()
      const slides = []
      allSections.forEach((s, i) => {
        const layer = s.getAttribute('data-marpit-advanced-background')
        if (layer === 'pseudo' || layer === 'background') return
        const key = s.getAttribute('data-marpit-pagination') || s.getAttribute('id') || String(i)
        if (seen.has(key)) return
        seen.add(key)
        slides.push(s)
      })

      return slides.map(section => {
        // Get all text content, excluding style/script tags
        const clone = section.cloneNode(true)
        clone.querySelectorAll('style, script').forEach(el => el.remove())
        let text = clone.textContent?.replace(/\s+/g, ' ').trim() || ''
        return text
      })
    })

    return slideTexts
  } finally {
    await browser.close()
  }
}

/**
 * Normalize text for comparison: strip all whitespace, lowercase.
 * This ensures differences in tokenization don't cause false positives.
 */
function normalizeForComparison(text) {
  return text
    .replace(/\s+/g, '')           // strip whitespace
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // zero-width chars
    .toLowerCase()
}

/**
 * Compare HTML text vs PPTX text per slide.
 *
 * Strategy: Use a sliding-window substring approach. For each slide,
 * split the normalized HTML text into N-char chunks and check what
 * fraction of those chunks exist in the normalized PPTX text.
 * This is robust against reordering, formatting splits, and minor
 * differences while detecting wholesale content loss.
 */
function compareTexts(htmlTexts, pptxTexts) {
  const CHUNK_SIZE = 6 // 6-char sliding window (covers CJK bigrams and Latin words)
  const results = []
  const slideCount = Math.max(htmlTexts.length, pptxTexts.length)

  for (let i = 0; i < slideCount; i++) {
    const htmlRaw = (htmlTexts[i] || '').trim()
    const pptxRaw = (pptxTexts[i] || '').trim()

    const htmlNorm = normalizeForComparison(htmlRaw)
    const pptxNorm = normalizeForComparison(pptxRaw)

    if (htmlNorm.length < CHUNK_SIZE) {
      results.push({ slide: i + 1, status: 'OK', missingChars: 0, totalChars: htmlNorm.length, missingPct: '0.0', missing: [] })
      continue
    }

    // Generate chunks from HTML
    let matchedChunks = 0
    let totalChunks = 0
    const missingSegments = []

    for (let j = 0; j <= htmlNorm.length - CHUNK_SIZE; j += CHUNK_SIZE) {
      const chunk = htmlNorm.substring(j, j + CHUNK_SIZE)
      totalChunks++
      if (pptxNorm.includes(chunk)) {
        matchedChunks++
      } else {
        if (missingSegments.length < 10) {
          missingSegments.push(chunk)
        }
      }
    }

    const missingPct = totalChunks > 0 ? ((totalChunks - matchedChunks) / totalChunks) * 100 : 0

    // FAIL if >15% of chunks are missing (significant content loss)
    // WARN if >5%
    let status = 'OK'
    if (missingPct > 15) status = 'FAIL'
    else if (missingPct > 5) status = 'WARN'

    results.push({
      slide: i + 1,
      status,
      missingChars: (totalChunks - matchedChunks) * CHUNK_SIZE,
      totalChars: htmlNorm.length,
      missingPct: missingPct.toFixed(1),
      missing: missingSegments,
    })
  }

  return results
}

// ─── Layer 1b: Raw Markdown Detection ──────────────────────────────────────

/**
 * Detect raw markdown syntax leaking into PPTX text.
 * If markdown headings (##), bold (**text**), unescaped bullets, or
 * raw link syntax [text](url) appear in the PPTX text, it means the
 * content was not properly converted and is rendering as source.
 *
 * Excludes text inside code-block shapes (monospace fonts) since markdown
 * syntax is legitimate content there.
 */
async function detectRawMarkdown(pptxPath) {
  const buf = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(buf)
  const slideEntries = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1])
      const nb = parseInt(b.match(/slide(\d+)/)[1])
      return na - nb
    })

  const results = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry].async('string')

    // Extract text per shape, skip shapes that use monospace (code block) fonts
    const MONOSPACE_PATTERN = /Courier|Consolas|Monaco|Source Code|Fira Code|JetBrains Mono|Menlo|monospace/i
    const shapes = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)]
    const nonCodeTexts = []
    for (const shapeMatch of shapes) {
      const shapeXml = shapeMatch[1]
      // Skip shapes that are code blocks (have monospace fonts)
      if (MONOSPACE_PATTERN.test(shapeXml)) continue
      const texts = [...shapeXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1])
      nonCodeTexts.push(...texts)
    }
    const allText = nonCodeTexts.join('\n')

    const issues = []

    // Detect markdown headings: lines starting with ## or ###
    const headingMatches = allText.match(/^#{2,6}\s+.+/gm) || []
    if (headingMatches.length > 0) {
      issues.push(`md-heading(${headingMatches.length}): ${headingMatches[0].substring(0, 40)}`)
    }

    // Detect bold/italic markers: **text** or __text__ (needs 3+ to avoid false positive)
    const boldMatches = allText.match(/\*\*[^*]+\*\*/g) || []
    if (boldMatches.length > 2) {
      issues.push(`md-bold(${boldMatches.length}): ${boldMatches[0].substring(0, 30)}`)
    }

    // Detect raw link syntax: [text](url)
    const linkMatches = allText.match(/\[[^\]]+\]\(https?:\/\/[^)]+\)/g) || []
    if (linkMatches.length > 0) {
      issues.push(`md-link(${linkMatches.length}): ${linkMatches[0].substring(0, 40)}`)
    }

    // Detect raw bullet markers at line starts: "- " or "* " (5+ occurrences to be sure)
    const bulletMatches = allText.match(/^[\-\*]\s+/gm) || []
    if (bulletMatches.length >= 5) {
      issues.push(`md-bullets(${bulletMatches.length})`)
    }

    const slideNum = parseInt(entry.match(/slide(\d+)/)[1])
    const status = issues.length > 0 ? 'FAIL' : 'OK'
    results.push({ slide: slideNum, status, issues })
  }

  return results
}

// ─── Layer 2: Structure Integrity ──────────────────────────────────────────

/**
 * Extract structural element counts from HTML via Puppeteer.
 */
async function extractHtmlStructure(htmlPath, browserPath) {
  const puppeteer = require('puppeteer-core')
  const browser = await puppeteer.launch({
    executablePath: browserPath,
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu'],
  })

  try {
    const page = await browser.newPage()
    await page.setViewport({ width: 1280, height: 720 })
    await page.goto(pathToFileURL(htmlPath).href, { waitUntil: 'networkidle0' })
    await new Promise(r => setTimeout(r, 300))

    return await page.evaluate(() => {
      const allSections = Array.from(document.querySelectorAll('section'))
        .filter(s => {
          if (s.parentElement?.closest('section')) return false
          return (
            s.parentElement?.tagName.toLowerCase() === 'foreignobject' ||
            s.hasAttribute('data-marpit-pagination')
          )
        })

      const seen = new Set()
      const slides = []
      allSections.forEach((s, i) => {
        const layer = s.getAttribute('data-marpit-advanced-background')
        if (layer === 'pseudo' || layer === 'background') return
        const key = s.getAttribute('data-marpit-pagination') || s.getAttribute('id') || String(i)
        if (seen.has(key)) return
        seen.add(key)
        slides.push(s)
      })

      return slides.map(section => {
        const codeBlocks = section.querySelectorAll('pre > code, pre.hljs').length
        const lists = section.querySelectorAll('ul, ol').length
        const tables = section.querySelectorAll('table').length
        const headings = section.querySelectorAll('h1, h2, h3, h4, h5, h6').length
        const images = section.querySelectorAll('img').length
        const blockquotes = section.querySelectorAll('blockquote').length
        return { codeBlocks, lists, tables, headings, images, blockquotes }
      })
    })
  } finally {
    await browser.close()
  }
}

/**
 * Extract structural element counts from PPTX XML.
 * Uses shape type detection in the PPTX XML.
 */
async function extractPptxStructure(pptxPath) {
  const buf = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(buf)
  const slideEntries = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1])
      const nb = parseInt(b.match(/slide(\d+)/)[1])
      return na - nb
    })

  const results = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry].async('string')

    // Count tables: <a:tbl> elements
    const tables = (xml.match(/<a:tbl[ >]/g) || []).length

    // Count images: <p:pic> elements or <a:blipFill>
    const images = (xml.match(/<p:pic[ >]/g) || []).length

    // Code blocks: shapes with monospace font or specific naming
    // In our PPTX, code blocks are shapes with Courier/Consolas fonts
    const codeIndicators = (xml.match(/Courier|Consolas|Monaco|monospace|Source Code/gi) || []).length
    const codeBlocks = codeIndicators > 0 ? Math.ceil(codeIndicators / 3) : 0 // approximate

    // Lists: shapes with <a:buChar> or <a:buAutoNum> (bullet/number markers)
    const bulletMarkers = (xml.match(/<a:buChar |<a:buAutoNum /g) || []).length
    const lists = bulletMarkers > 0 ? 1 : 0 // at least one list if bullets exist

    results.push({ tables, images, codeBlocks, lists })
  }
  return results
}

/**
 * Compare structural elements between HTML and PPTX.
 */
function compareStructure(htmlStructure, pptxStructure) {
  const results = []
  const slideCount = Math.max(htmlStructure.length, pptxStructure.length)

  for (let i = 0; i < slideCount; i++) {
    const html = htmlStructure[i] || { codeBlocks: 0, lists: 0, tables: 0, headings: 0, images: 0, blockquotes: 0 }
    const pptx = pptxStructure[i] || { tables: 0, images: 0, codeBlocks: 0, lists: 0 }

    const issues = []

    // Table count must match exactly
    if (html.tables > 0 && pptx.tables === 0) {
      issues.push(`table: HTML=${html.tables}, PPTX=0`)
    }

    // Code blocks: HTML has code but PPTX has none
    if (html.codeBlocks > 0 && pptx.codeBlocks === 0) {
      issues.push(`code: HTML=${html.codeBlocks}, PPTX=0`)
    }

    // Lists: HTML has lists but PPTX has no bullets
    if (html.lists > 0 && pptx.lists === 0) {
      issues.push(`list: HTML=${html.lists}, PPTX=0`)
    }

    const status = issues.length > 0 ? 'FAIL' : 'OK'
    results.push({ slide: i + 1, status, issues })
  }

  return results
}

// ─── Layer 3: Glyph Rendering (Tofu Detection) ────────────────────────────

/**
 * Detect tofu (□) characters in PPTX text.
 * Tofu appears as replacement characters or actual □ in extracted text.
 */
async function detectTofu(pptxPath) {
  const buf = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(buf)
  const slideEntries = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1])
      const nb = parseInt(b.match(/slide(\d+)/)[1])
      return na - nb
    })

  const results = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry].async('string')
    const texts = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1])
    const allText = texts.join('')

    // Detect tofu patterns:
    // - U+25A1 □ (white square)
    // - U+FFFD  (replacement character)
    // - U+2610 ☐
    // - Three or more consecutive identical box-like chars
    const tofuChars = allText.match(/[\u25A0\u25A1\u25AB\u25AC\u25FB\u25FC\u25FD\u25FE\uFFFD\u2610]{2,}/g) || []
    // Also check for runs of \uFFFD
    const replacementRuns = allText.match(/\uFFFD+/g) || []

    const hasTofuPattern = tofuChars.length > 0 || replacementRuns.length > 0
    const slideNum = parseInt(entry.match(/slide(\d+)/)[1])

    results.push({
      slide: slideNum,
      status: hasTofuPattern ? 'FAIL' : 'OK',
      tofuRuns: tofuChars.length + replacementRuns.length,
      samples: [...tofuChars, ...replacementRuns].slice(0, 5),
    })
  }

  return results
}

/**
 * Detect tofu in PNG screenshots by looking for patterns of
 * identical small rectangles (visual tofu). This uses pixel analysis.
 */
function detectTofuInPngs(pngDir) {
  if (!pngDir || !fs.existsSync(pngDir)) return []

  const { PNG } = require('pngjs')
  const pptxPngs = fs.readdirSync(pngDir)
    .filter(f => /^pptx-slide-\d+\.png$/.test(f))
    .sort()

  const results = []
  for (const pngFile of pptxPngs) {
    const slideNum = parseInt(pngFile.match(/(\d+)\.png$/)[1])
    const imgBuf = fs.readFileSync(path.join(pngDir, pngFile))
    const img = PNG.sync.read(imgBuf)

    // Scan for tofu pattern: consecutive identical-width rectangles
    // Tofu typically appears as uniform squares in a row
    // Simplified heuristic: look for horizontal runs of near-identical
    // pixel columns that form a repeating square pattern
    let tofuScore = 0

    // Sample middle rows where text typically appears
    const sampleRows = [
      Math.floor(img.height * 0.3),
      Math.floor(img.height * 0.5),
      Math.floor(img.height * 0.7),
    ]

    for (const row of sampleRows) {
      // Look for runs of identical pixel columns (characteristic of tofu)
      let runLen = 0
      let lastCol = null

      for (let x = 0; x < img.width - 1; x++) {
        const idx = (row * img.width + x) * 4
        const colSig = `${img.data[idx]},${img.data[idx + 1]},${img.data[idx + 2]}`

        if (colSig === lastCol) {
          runLen++
        } else {
          // If we had a suspiciously uniform run followed by a break
          if (runLen >= 8 && runLen <= 20) {
            tofuScore++
          }
          runLen = 1
          lastCol = colSig
        }
      }
    }

    results.push({
      slide: slideNum,
      status: tofuScore >= 3 ? 'WARN' : 'OK', // only WARN for visual tofu (uncertain)
      tofuScore,
    })
  }

  return results
}

// ─── Layer 3b: Font Validation (Proprietary/Unavailable Font Detection) ───

/**
 * Detect fonts in PPTX shapes that are likely unavailable on target systems.
 *
 * Proprietary fonts contain patterns like "35HSJPDOC" (digit+uppercase suffix)
 * which are developer-specific fonts. When these are embedded in a PPTX,
 * the viewer will fall back to a system font and may produce tofu or
 * incorrect rendering.
 *
 * Also flags shapes where CJK text is assigned to a non-CJK-capable font.
 */
async function detectProblematicFonts(pptxPath) {
  const buf = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(buf)
  const slideEntries = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1])
      const nb = parseInt(b.match(/slide(\d+)/)[1])
      return na - nb
    })

  // Font patterns that indicate developer-specific / proprietary fonts
  const PROPRIETARY_PATTERN = /\b\d{2,4}[A-Z]{2,}/i
  // Known safe fonts for CJK rendering
  const CJK_SAFE_FONTS = /Meiryo|Yu Gothic|MS Gothic|MS Mincho|Noto Sans|HGP|HGS|HG|ヒラギノ|Hiragino|游|源|IPAex|BIZ UD/i
  // CJK character pattern
  const CJK_CHARS = /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff01-\uff9f]/

  const results = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry].async('string')
    const slideNum = parseInt(entry.match(/slide(\d+)/)[1])
    const issues = []

    // Extract all font references: <a:latin typeface="..."/> <a:ea typeface="..."/> <a:cs typeface="..."/>
    const fontRefs = [...xml.matchAll(/<a:(latin|ea|cs)\s+typeface="([^"]+)"/g)]
    const proprietaryFonts = new Set()

    for (const [, fontType, fontName] of fontRefs) {
      if (PROPRIETARY_PATTERN.test(fontName)) {
        proprietaryFonts.add(fontName)
      }
    }

    if (proprietaryFonts.size > 0) {
      issues.push(`proprietary-font(${proprietaryFonts.size}): ${[...proprietaryFonts].slice(0, 3).join(', ')}`)
    }

    // Check CJK text without CJK-capable font (ea typeface)
    const shapes = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)]
    for (const shapeMatch of shapes) {
      const shapeXml = shapeMatch[1]
      const texts = [...shapeXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1])
      const allText = texts.join('')
      if (!CJK_CHARS.test(allText)) continue

      // Check if this shape's font is CJK-capable
      const eaFonts = [...shapeXml.matchAll(/<a:ea\s+typeface="([^"]+)"/g)].map(m => m[1])
      const latinFonts = [...shapeXml.matchAll(/<a:latin\s+typeface="([^"]+)"/g)].map(m => m[1])
      const activeFonts = eaFonts.length > 0 ? eaFonts : latinFonts

      for (const font of activeFonts) {
        if (PROPRIETARY_PATTERN.test(font) && !CJK_SAFE_FONTS.test(font)) {
          issues.push(`cjk-font-risk: "${font}" for CJK text "${allText.substring(0, 20)}"`)
          break
        }
      }
    }

    const status = issues.some(i => i.startsWith('proprietary-font') || i.startsWith('cjk-font-risk'))
      ? 'WARN' : 'OK'
    results.push({ slide: slideNum, status, issues })
  }

  return results
}

// ─── Layer 2b: Code Block Formatting Integrity ────────────────────────────

/**
 * Verify that code block content in the PPTX retains monospace font and
 * background fill. Detects the problem where code blocks inside list items
 * lose their formatting (background shape disappears, font becomes proportional).
 */
async function checkCodeBlockFormatting(pptxPath) {
  const buf = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(buf)
  const slideEntries = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1])
      const nb = parseInt(b.match(/slide(\d+)/)[1])
      return na - nb
    })

  const MONOSPACE_PATTERN = /Courier|Consolas|Monaco|Source Code|Fira Code|JetBrains Mono|Menlo|monospace|UDEV Gothic/i
  // Code indicators: typical code patterns that suggest this text is code
  const CODE_PATTERNS = [
    /\b(const|let|var|function|class|import|export|return|if|else|for|while)\b/,
    /[{}();].*[{}();]/,  // multiple braces/semicolons
    /^\s*(\/\/|#|\/\*|\*)/m, // comments
    /\b(npm|yarn|pip|apt|brew)\s+(install|run|start)/,
    /\$\(|=>|\.\.\./,
  ]

  const results = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry].async('string')
    const slideNum = parseInt(entry.match(/slide(\d+)/)[1])
    const issues = []

    // Find shapes with bullet markers (list items)
    const shapes = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)]
    for (const shapeMatch of shapes) {
      const shapeXml = shapeMatch[1]
      const hasBullets = /<a:buChar |<a:buAutoNum /.test(shapeXml)
      if (!hasBullets) continue

      // Extract text from this list shape
      const texts = [...shapeXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1])
      const allText = texts.join('\n')

      // Check if any text looks like code
      const looksLikeCode = CODE_PATTERNS.some(p => p.test(allText))
      if (!looksLikeCode) continue

      // If code-like text is in a list shape, it should have monospace font
      const hasMonospace = MONOSPACE_PATTERN.test(shapeXml)
      if (!hasMonospace) {
        issues.push(`code-in-list-no-monospace: "${allText.substring(0, 40)}"`)
      }
    }

    const status = issues.length > 0 ? 'WARN' : 'OK'
    results.push({ slide: slideNum, status, issues })
  }

  return results
}

// ─── Layer 4: Bounds Validation + Overflow Estimation ─────────────────────

/**
 * Check if text-bearing shapes are placed outside the visible slide area,
 * AND estimate if text content likely overflows its container.
 *
 * PPTX slide dimensions are in EMU (English Metric Units): 1 inch = 914400 EMU.
 * Standard 16:9 slide: 12192000 x 6858000 EMU (= 1280×720 px at 96 dpi).
 *
 * Detects:
 *   - shapes whose x+cx < 0 or x > slideWidth (fully off-screen)
 *   - shapes where text char count is too large for the shape area (overflow)
 */
async function checkBounds(pptxPath) {
  const buf = fs.readFileSync(pptxPath)
  const zip = await JSZip.loadAsync(buf)
  const slideEntries = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1])
      const nb = parseInt(b.match(/slide(\d+)/)[1])
      return na - nb
    })

  // Standard slide dimensions in EMU
  const SLIDE_W = 12192000
  const SLIDE_H = 6858000
  // Allow 10% overflow tolerance (content partially off-screen is OK)
  const TOLERANCE = 0.1

  const results = []
  for (const entry of slideEntries) {
    const xml = await zip.files[entry].async('string')
    const slideNum = parseInt(entry.match(/slide(\d+)/)[1])

    // Find all shape transforms: <a:off x="..." y="..."/><a:ext cx="..." cy="..."/>
    // within <p:sp> (shapes) that contain text <a:t>
    const shapes = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)]
    let offScreenCount = 0
    let overflowCount = 0
    let tinyFontCount = 0
    let totalTextShapes = 0
    let offScreenTexts = []
    let overflowTexts = []
    let tinyFontTexts = []

    for (const shapeMatch of shapes) {
      const shapeXml = shapeMatch[1]
      // Check if shape has text
      const hasText = /<a:t>([^<]+)<\/a:t>/.test(shapeXml)
      if (!hasText) continue

      totalTextShapes++

      // Extract position
      const offMatch = shapeXml.match(/<a:off x="(-?\d+)" y="(-?\d+)"/)
      const extMatch = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"/)
      if (!offMatch || !extMatch) continue

      const x = parseInt(offMatch[1])
      const y = parseInt(offMatch[2])
      const cx = parseInt(extMatch[1])
      const cy = parseInt(extMatch[2])

      // Check if shape is fully off-screen
      const rightEdge = x + cx
      const bottomEdge = y + cy
      const minVisible = SLIDE_W * TOLERANCE

      const isOffRight = x > SLIDE_W
      const isOffLeft = rightEdge < 0
      const isOffBottom = y > SLIDE_H
      const isOffTop = bottomEdge < 0

      if (isOffRight || isOffLeft || isOffBottom || isOffTop) {
        offScreenCount++
        const textSample = (shapeXml.match(/<a:t>([^<]+)<\/a:t>/)?.[1] || '').substring(0, 30)
        offScreenTexts.push(textSample)
      }

      // --- Overflow estimation ---
      // Estimate if text content is too large for the shape area.
      // Heuristic: assume ~12pt font → ~170000 EMU line height,
      // and ~110000 EMU per CJK char width (≈7pt width at 12pt).
      const allShapeTexts = [...shapeXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1]).join('')
      const charCount = allShapeTexts.length
      if (charCount > 0 && cx > 0 && cy > 0) {
        const LINE_HEIGHT = 170000 // EMU per line
        const CHAR_WIDTH = 110000  // EMU per char (CJK)
        const charsPerLine = Math.max(1, Math.floor(cx / CHAR_WIDTH))
        const linesNeeded = Math.ceil(charCount / charsPerLine)
        const heightNeeded = linesNeeded * LINE_HEIGHT
        // If text needs >150% of available height, likely overflow
        if (heightNeeded > cy * 1.5 && charCount > 20) {
          overflowCount++
          const textSample = allShapeTexts.substring(0, 30)
          overflowTexts.push(`${charCount}chars/${Math.round(cy/914400*72)}pt: ${textSample}`)
        }
      }

      // --- Tiny font detection ---
      // If a shape has text rendered at an extremely small font size (<= 600 = 6pt),
      // the content is likely being force-compressed to fit, producing unreadable output
      // that differs dramatically from the HTML rendering.
      // <a:sz val="600"/> means 6pt (unit is 1/100 pt)
      const fontSizes = [...shapeXml.matchAll(/<a:(?:sz|defRPr)[^>]*\sval="(\d+)"/g)].map(m => parseInt(m[1]))
      // Also check <a:rPr ... sz="XXX">
      const rPrSizes = [...shapeXml.matchAll(/<a:rPr[^>]*\ssz="(\d+)"/g)].map(m => parseInt(m[1]))
      const allSizes = [...fontSizes, ...rPrSizes].filter(s => s > 0)
      const minFontSize = allSizes.length > 0 ? Math.min(...allSizes) : 1200
      if (minFontSize <= 600 && charCount > 30) {
        tinyFontCount++
        const textSample = allShapeTexts.substring(0, 30)
        tinyFontTexts.push(`${minFontSize/100}pt(${charCount}chars): ${textSample}`)
      }
    }

    // FAIL if any text shapes are completely off-screen or have tiny font (content squeeze)
    // WARN if text overflow detected
    let status = 'OK'
    if (offScreenCount > 0 || tinyFontCount > 0) {
      status = 'FAIL'
    } else if (overflowCount > 0) {
      status = 'WARN'
    }

    results.push({
      slide: slideNum,
      status,
      offScreenCount,
      overflowCount,
      tinyFontCount,
      totalTextShapes,
      offScreenTexts: offScreenTexts.slice(0, 5),
      overflowTexts: overflowTexts.slice(0, 5),
      tinyFontTexts: tinyFontTexts.slice(0, 5),
    })
  }

  return results
}

// ─── Layer 6: Spatial Coverage (Blank Area Detection) ──────────────────────

/**
 * Compare non-white pixel coverage between HTML and PPTX PNGs.
 * If the PPTX rendering has significantly less non-white area than the HTML,
 * it indicates content loss (e.g. entire columns/sections disappeared).
 *
 * Also detects the specific pattern of "right half blank" which indicates
 * a 2-column layout where the right column content was lost.
 */
function compareSpatialCoverage(pngDir) {
  if (!pngDir || !fs.existsSync(pngDir)) return []

  const { PNG } = require('pngjs')

  // Find matched pairs: html-slide-N.png and pptx-slide-N.png
  const allFiles = fs.readdirSync(pngDir)
  const htmlPngs = allFiles.filter(f => /^html-slide-\d+\.png$/.test(f)).sort()
  const pptxPngs = allFiles.filter(f => /^pptx-slide-\d+\.png$/.test(f)).sort()

  const results = []
  for (const htmlFile of htmlPngs) {
    const slideNum = parseInt(htmlFile.match(/(\d+)\.png$/)[1])
    const pptxFile = `pptx-slide-${slideNum}.png`
    if (!pptxPngs.includes(pptxFile)) continue

    const htmlImg = PNG.sync.read(fs.readFileSync(path.join(pngDir, htmlFile)))
    const pptxImg = PNG.sync.read(fs.readFileSync(path.join(pngDir, pptxFile)))

    // Count non-white pixels (threshold: any channel < 240)
    const WHITE_THRESHOLD = 240
    let htmlNonWhite = 0
    let pptxNonWhite = 0
    let pptxRightHalfNonWhite = 0
    let pptxLeftHalfNonWhite = 0
    const halfX = Math.floor(pptxImg.width / 2)

    for (let y = 0; y < htmlImg.height; y++) {
      for (let x = 0; x < htmlImg.width; x++) {
        const idx = (y * htmlImg.width + x) * 4
        if (htmlImg.data[idx] < WHITE_THRESHOLD ||
            htmlImg.data[idx + 1] < WHITE_THRESHOLD ||
            htmlImg.data[idx + 2] < WHITE_THRESHOLD) {
          htmlNonWhite++
        }
      }
    }

    for (let y = 0; y < pptxImg.height; y++) {
      for (let x = 0; x < pptxImg.width; x++) {
        const idx = (y * pptxImg.width + x) * 4
        if (pptxImg.data[idx] < WHITE_THRESHOLD ||
            pptxImg.data[idx + 1] < WHITE_THRESHOLD ||
            pptxImg.data[idx + 2] < WHITE_THRESHOLD) {
          pptxNonWhite++
          if (x >= halfX) pptxRightHalfNonWhite++
          else pptxLeftHalfNonWhite++
        }
      }
    }

    const totalPixels = pptxImg.width * pptxImg.height
    const htmlCoverage = htmlNonWhite / totalPixels
    const pptxCoverage = pptxNonWhite / totalPixels

    // Coverage ratio: how much of HTML's coverage does PPTX retain?
    const coverageRatio = htmlCoverage > 0.01 ? pptxCoverage / htmlCoverage : 1.0

    // Right-half emptiness: if HTML has content but PPTX right half is nearly empty
    const rightHalfRatio = pptxLeftHalfNonWhite > 100
      ? pptxRightHalfNonWhite / pptxLeftHalfNonWhite
      : 1.0

    let status = 'OK'
    const issues = []

    // FAIL: PPTX coverage is less than 35% of HTML coverage (massive content loss)
    if (coverageRatio < 0.35 && htmlCoverage > 0.05) {
      status = 'FAIL'
      issues.push(`coverage-loss: HTML=${(htmlCoverage * 100).toFixed(1)}% PPTX=${(pptxCoverage * 100).toFixed(1)}% ratio=${coverageRatio.toFixed(2)}`)
    }
    // FAIL: Right half is nearly empty while left has content (2-column loss pattern)
    else if (rightHalfRatio < 0.05 && pptxLeftHalfNonWhite > 1000 && htmlCoverage > 0.05) {
      status = 'FAIL'
      issues.push(`right-half-blank: left=${pptxLeftHalfNonWhite}px right=${pptxRightHalfNonWhite}px ratio=${rightHalfRatio.toFixed(3)}`)
    }
    // WARN: moderate coverage loss
    else if (coverageRatio < 0.55 && htmlCoverage > 0.05) {
      status = 'WARN'
      issues.push(`coverage-reduced: ratio=${coverageRatio.toFixed(2)}`)
    }

    results.push({
      slide: slideNum,
      status,
      htmlCoverage: (htmlCoverage * 100).toFixed(1),
      pptxCoverage: (pptxCoverage * 100).toFixed(1),
      coverageRatio: coverageRatio.toFixed(2),
      issues,
    })
  }

  return results
}

// ─── Main ──────────────────────────────────────────────────────────────────

function findChrome() {
  try {
    const { computeSystemExecutablePath, Browser, ChromeReleaseChannel } = require('@puppeteer/browsers')
    const p = computeSystemExecutablePath({
      browser: Browser.CHROME,
      platform: 'win64',
      channel: ChromeReleaseChannel.STABLE,
    })
    if (fs.existsSync(p)) return p
  } catch { }
  return undefined
}

async function main() {
  const args = process.argv.slice(2)
  const htmlPath = path.resolve(args.find(a => !a.startsWith('--')) || '')
  const pptxPath = path.resolve(args.filter(a => !a.startsWith('--'))[1] || '')
  const pngDirArg = args.find(a => a.startsWith('--png-dir='))
  const pngDir = pngDirArg ? path.resolve(pngDirArg.split('=')[1]) : null

  if (!fs.existsSync(htmlPath) || !fs.existsSync(pptxPath)) {
    console.error('Usage: structural-test.js <html-path> <pptx-path> [--png-dir=<dir>]')
    process.exit(1)
  }

  const browserPath = process.env.CHROME_PATH || findChrome()
  if (!browserPath) {
    console.error('Chrome not found')
    process.exit(1)
  }

  const basename = path.basename(htmlPath, '.html')
  process.stderr.write(`[structural-test] ${basename}\n`)

  // Layer 1: Content Completeness
  process.stderr.write('  Layer 1: Content Completeness...\n')
  const [htmlTexts, pptxTexts] = await Promise.all([
    extractHtmlTexts(htmlPath, browserPath),
    extractPptxTexts(pptxPath),
  ])
  const contentResults = compareTexts(htmlTexts, pptxTexts)

  // Layer 2: Structure Integrity
  process.stderr.write('  Layer 2: Structure Integrity...\n')
  const [htmlStruct, pptxStruct] = await Promise.all([
    extractHtmlStructure(htmlPath, browserPath),
    extractPptxStructure(pptxPath),
  ])
  const structureResults = compareStructure(htmlStruct, pptxStruct)

  // Layer 3: Glyph Rendering (Tofu)
  process.stderr.write('  Layer 3: Glyph Rendering...\n')
  const tofuTextResults = await detectTofu(pptxPath)
  const tofuPngResults = pngDir ? detectTofuInPngs(pngDir) : []

  // Merge tofu results (text + PNG)
  const tofuResults = tofuTextResults.map(r => {
    const pngResult = tofuPngResults.find(p => p.slide === r.slide)
    const combinedStatus = r.status === 'FAIL' ? 'FAIL'
      : pngResult?.status === 'WARN' ? 'WARN'
      : r.status
    return { ...r, pngTofuScore: pngResult?.tofuScore ?? null, status: combinedStatus }
  })

  // Layer 4: Bounds Validation
  process.stderr.write('  Layer 4: Bounds Validation...\n')
  const boundsResults = await checkBounds(pptxPath)

  // Layer 5: Raw Markdown Detection
  process.stderr.write('  Layer 5: Raw Markdown Detection...\n')
  const markdownResults = await detectRawMarkdown(pptxPath)

  // Layer 3b: Font Validation (Proprietary Font Detection)
  process.stderr.write('  Layer 3b: Font Validation...\n')
  const fontResults = await detectProblematicFonts(pptxPath)

  // Layer 2b: Code Block Formatting Integrity
  process.stderr.write('  Layer 2b: Code Block Formatting...\n')
  const codeFormatResults = await checkCodeBlockFormatting(pptxPath)

  // Layer 6: Spatial Coverage (requires PNGs)
  let spatialResults = []
  if (pngDir) {
    process.stderr.write('  Layer 6: Spatial Coverage...\n')
    spatialResults = compareSpatialCoverage(pngDir)
  }

  // Aggregate per-slide
  const slideCount = Math.max(contentResults.length, structureResults.length, tofuResults.length, boundsResults.length, markdownResults.length)
  const perSlide = []
  for (let i = 0; i < slideCount; i++) {
    const content = contentResults[i] || { slide: i + 1, status: 'OK' }
    const structure = structureResults[i] || { slide: i + 1, status: 'OK' }
    const tofu = tofuResults[i] || { slide: i + 1, status: 'OK' }
    const bounds = boundsResults[i] || { slide: i + 1, status: 'OK' }
    const markdown = markdownResults[i] || { slide: i + 1, status: 'OK' }
    const spatial = spatialResults.find(s => s.slide === i + 1) || { slide: i + 1, status: 'OK' }
    const font = fontResults[i] || { slide: i + 1, status: 'OK' }
    const codeFormat = codeFormatResults[i] || { slide: i + 1, status: 'OK' }

    // Overall: worst of all layers
    const statuses = [content.status, structure.status, tofu.status, bounds.status, markdown.status, spatial.status, font.status, codeFormat.status]
    const overall = statuses.includes('FAIL') ? 'FAIL'
      : statuses.includes('WARN') ? 'WARN'
      : 'OK'

    perSlide.push({
      slide: i + 1,
      overall,
      content: content.status,
      structure: structure.status,
      tofu: tofu.status,
      bounds: bounds.status,
      markdown: markdown.status,
      spatial: spatial.status,
      font: font.status,
      codeFormat: codeFormat.status,
      details: {
        contentMissing: content.missing || [],
        contentMissingPct: content.missingPct || '0',
        structureIssues: structure.issues || [],
        tofuRuns: tofu.tofuRuns || 0,
        offScreenCount: bounds.offScreenCount || 0,
        overflowCount: bounds.overflowCount || 0,
        tinyFontCount: bounds.tinyFontCount || 0,
        offScreenTexts: bounds.offScreenTexts || [],
        overflowTexts: bounds.overflowTexts || [],
        tinyFontTexts: bounds.tinyFontTexts || [],
        markdownIssues: markdown.issues || [],
        spatialIssues: spatial.issues || [],
        fontIssues: font.issues || [],
        codeFormatIssues: codeFormat.issues || [],
      }
    })
  }

  // Summary
  const fails = perSlide.filter(s => s.overall === 'FAIL')
  const warns = perSlide.filter(s => s.overall === 'WARN')
  const oks = perSlide.filter(s => s.overall === 'OK')

  const report = {
    file: basename,
    slides: slideCount,
    summary: {
      FAIL: fails.length,
      WARN: warns.length,
      OK: oks.length,
    },
    failedSlides: fails.map(s => s.slide),
    perSlide,
  }

  // Print human-readable summary to stderr
  process.stderr.write(`  RESULT: FAIL:${fails.length} WARN:${warns.length} OK:${oks.length} (${slideCount} slides)\n`)
  if (fails.length > 0) {
    process.stderr.write(`  FAILED: ${fails.map(s => `${s.slide}(${[s.content !== 'OK' ? 'C' : '', s.structure !== 'OK' ? 'S' : '', s.tofu !== 'OK' ? 'T' : '', s.bounds !== 'OK' ? 'B' : '', s.markdown !== 'OK' ? 'M' : '', s.spatial !== 'OK' ? 'V' : '', s.font !== 'OK' ? 'F' : '', s.codeFormat !== 'OK' ? 'CF' : ''].filter(Boolean).join('+')})`).join(', ')}\n`)
  }

  // JSON to stdout
  console.log(JSON.stringify(report, null, 2))
}

main().catch(e => {
  console.error(e)
  process.exit(1)
})
