import { fileURLToPath } from 'node:url'
import PptxGenJS from 'pptxgenjs'
import type {
  SlideData,
  SlideElement,
  TextRun,
  ListItem,
  TextStyle,
} from './types'
import {
  rgbToHex,
  compositeOver,
  cleanFontFamily,
  pxToInches,
  pxToPoints,
  isTransparent,
  sanitizeText,
} from './utils'

/**
 * Proportional scale factor applied to each table column width when building
 * PptxGenJS output.  DirectWrite (PowerPoint) renders fonts slightly wider than
 * Chrome's Skia engine, causing table header text that fits in the browser to
 * wrap in PPTX.  A 5% increase covers the observed ~3% variance across all
 * measured header cell lengths and font sizes (see ADR-26).
 */
const DIRECTWRITE_COL_WIDTH_FACTOR = 1.05
/**
 * Slide canvas width in CSS pixels (1280×720 WIDEP layout at 96 dpi).
 * Used as the maximum right-edge boundary for column width overflow guard.
 */
const SLIDE_CANVAS_W_PX = 1280

/**
 * Maps a CSS border-style string to the PptxGenJS dashType string.
 * Returns undefined for solid (default, no dashType needed) and for CSS
 * styles that have no PptxGenJS equivalent (e.g. 'double').
 */
function cssBorderStyleToDash(
  borderStyle: string | undefined,
): 'dash' | 'sysDot' | undefined {
  if (!borderStyle || borderStyle === 'solid') return undefined
  if (borderStyle === 'dashed') return 'dash'
  if (borderStyle === 'dotted') return 'sysDot'
  // 'double' and other CSS styles have no PptxGenJS dashType equivalent
  return undefined
}

/** Resolve a URL (data:, file:, or http) into PptxGenJS image source props. */
function resolveImageSource(url: string): { data?: string; path?: string } {
  if (url.startsWith('data:')) return { data: url }
  if (url.startsWith('file:')) return { path: fileURLToPath(url) }
  return { path: url }
}

/**
 * Convert CSS line-height and font-size into a PPTX lineSpacingMultiple value.
 *
 * Both values are in px (from getComputedStyle).  Returns undefined when the
 * ratio is outside a sensible range so PptxGenJS uses its own default.
 */
function computeLineSpacing(style: TextStyle): number | undefined {
  const { lineHeight, fontSize } = style
  if (!lineHeight || !fontSize || lineHeight <= 0 || fontSize <= 0)
    return undefined
  const m = lineHeight / fontSize
  if (m < 0.5 || m > 4) return undefined
  return Math.round(m * 100) / 100
}

/**
 * Convert CSS letter-spacing (px) into a PptxGenJS charSpacing value (points).
 * Returns undefined when the value is negligible.
 */
function computeCharSpacing(style: TextStyle): number | undefined {
  const ls = style.letterSpacing
  if (!ls || Math.abs(ls) < 0.1) return undefined
  return Math.round(pxToPoints(ls) * 100) / 100
}

/**
 * Convert CSS padding values (px) into a PptxGenJS margin (inset) tuple.
 * Returns 0 (no inset) when the style has no padding fields.
 *
 * IMPORTANT: PptxGenJS maps the 4-element margin array to OOXML bodyPr as
 *   [0]→lIns  [1]→rIns  [2]→bIns  [3]→tIns
 * This is NOT the CSS shorthand order (top/right/bottom/left).
 * We therefore return [left, right, bottom, top] so the OOXML values are correct.
 */
function computeTextInset(
  style: TextStyle,
): [number, number, number, number] | 0 {
  const pt = (style.paddingTop ?? 0) * 0.75
  const pr = (style.paddingRight ?? 0) * 0.75
  const pb = (style.paddingBottom ?? 0) * 0.75
  const pl = (style.paddingLeft ?? 0) * 0.75
  // PptxGenJS margin[0]→lIns, [1]→rIns, [2]→bIns, [3]→tIns
  return pt || pr || pb || pl ? [pl, pr, pb, pt] : 0
}

function createSlideNumberProps(
  slideW: number,
  slideH: number,
): PptxGenJS.SlideNumberProps {
  const bottomInset = slideH * (30 / 720)
  const rightInset = slideW * (40 / 1280)
  const boxHeight = slideH * (24 / 720)
  const fontSize = slideH * (18 / 720)

  return {
    x: 0,
    y: pxToInches(Math.max(0, slideH - bottomInset - boxHeight)),
    w: pxToInches(Math.max(48, slideW - rightInset)),
    h: pxToInches(Math.max(1, boxHeight)),
    align: 'right',
    color: '777777',
    fontSize: Math.round(pxToPoints(fontSize) * 100) / 100,
    margin: 0,
  }
}

const DEFAULT_BULLET_INDENT_PT = 27

function createListBulletOption(
  item: ListItem,
  ordered: boolean,
  continuation = false,
): boolean | Record<string, any> {
  const extraIndent = item.leadingOffset ? pxToPoints(item.leadingOffset) : 0
  const indent =
    extraIndent > 0
      ? Math.round((DEFAULT_BULLET_INDENT_PT + extraIndent) * 100) / 100
      : undefined

  if (continuation) {
    return {
      characterCode: '200B',
      ...(indent !== undefined ? { indent } : {}),
    }
  }

  if (ordered) {
    return {
      type: 'number',
      style: 'arabicPeriod',
      ...(indent !== undefined ? { indent } : {}),
    }
  }

  return indent !== undefined ? { indent } : true
}

/**
 * Build a PptxGenJS presentation from structured slide data extracted by the
 * DOM walker.
 */
export function buildPptx(slides: SlideData[]): PptxGenJS {
  const pptx = new PptxGenJS()

  const slideW = slides[0]?.width ?? 1280
  const slideH = slides[0]?.height ?? 720

  pptx.defineLayout({
    name: 'MARP',
    width: pxToInches(slideW),
    height: pxToInches(slideH),
  })
  pptx.layout = 'MARP'

  const useSlideNumbers = slides.some((slide) => slide.sourceHasPagination)

  for (const slideData of slides) {
    const slide = pptx.addSlide()
    if (useSlideNumbers) {
      slide.slideNumber = {
        ...createSlideNumberProps(slideW, slideH),
      }
    }

    // Slide background color (used when no full-slide background image exists)
    const bgColor = isTransparent(slideData.background)
      ? 'FFFFFF'
      : rgbToHex(slideData.background)

    const bgImages = slideData.backgroundImages ?? []

    // Determine if the first background image is a full-slide cover without a
    // CSS filter — if so, use it as the PPTX slide background property (which
    // is the proper way to set a slide background in OOXML and gives the best
    // editing experience in PowerPoint).
    const firstBg = bgImages[0]
    const isFullSlide =
      firstBg &&
      !firstBg.cssFilter &&
      firstBg.x <= 1 &&
      firstBg.y <= 1 &&
      Math.abs(firstBg.width - slideData.width) <= 2 &&
      Math.abs(firstBg.height - slideData.height) <= 2

    if (isFullSlide && bgImages.length === 1) {
      // Single full-slide background without filter → use slide.background
      slide.background = resolveImageSource(firstBg.url)
    } else {
      // Multiple backgrounds or partial/filtered backgrounds → solid fill +
      // overlay each background image as a positioned shape.
      slide.background = { fill: bgColor }

      for (const bg of bgImages) {
        const x = pxToInches(bg.x)
        const y = pxToInches(bg.y)
        const w = pxToInches(bg.width)
        const h = pxToInches(bg.height)
        const imgOpts: PptxGenJS.ImageProps = {
          x,
          y,
          w,
          h,
          ...resolveImageSource(bg.url),
        }
        slide.addImage(imgOpts)
      }
    }

    // Place elements at absolute coordinates
    for (const el of slideData.elements) {
      // Detect image-backed dark slides: bg image(s) present AND CSS bg-color
      // fell back to white (visual bg is provided by the image, not CSS).
      // CSS gradient placeholders (fromCssFallback=true) must NOT count as
      // real dark background images — they represent light gradients captured
      // pixel-for-pixel and cannot make the slide visually dark.
      // Only user-supplied ![bg] images (fromCssFallback absent) can produce
      // a dark background that would make light inline-code highlights invisible.
      const bgImages = slideData.backgroundImages ?? []
      const cssIsFallbackWhite =
        !slideData.background ||
        rgbToHex(slideData.background).toUpperCase() === 'FFFFFF'
      const visualBgMayBeDark =
        bgImages.some(
          (bg) =>
            bg.url !== '' &&
            !bg.fromCssFallback &&
            // Partial-width split backgrounds (e.g. `![bg right:30%]`) leave
            // the text area on a white background — treat them as non-dark.
            // Only full-slide images (width ≥ 80% of slide) may make the
            // visual background dark enough to suppress light code highlights.
            bg.width >= slideData.width * 0.8,
        ) && cssIsFallbackWhite
      placeElement(
        slide,
        el,
        slideData.width,
        slideData.height,
        slideData.background ?? 'rgb(255, 255, 255)',
        visualBgMayBeDark,
      )
    }

    // Presenter notes
    if (slideData.notes) {
      slide.addNotes(slideData.notes)
    }
  }

  return pptx
}

// Text element types whose height should be clamped to slide bounds.
// Font rendering differences between browser and PPTX can cause text
// boxes near the slide bottom to extend beyond the visible area.
// Images and containers are intentionally excluded — overflow can be
// valid (e.g. bleed images, split-layout backgrounds).
const TEXT_ELEMENT_TYPES = new Set([
  'heading',
  'paragraph',
  'list',
  'blockquote',
  'code',
  'table',
  'header',
  'footer',
])

/**
 * Compute the PPTX highlight hex string for a text run's backgroundColor.
 *
 * Strategy:
 *  1. Composite the (possibly semi-transparent) backgroundColor over the actual
 *     slide background color so we get an opaque approximation that matches what
 *     the browser renders.  Using the real slide bg (instead of always white) is
 *     critical for dark-background slides: rgba(0.12) over dark → slightly
 *     lighter dark, not near-white.
 *  2. Suppress when the composited color is too close to the slide bg (max channel
 *     delta < 15) — the highlight would be invisible anyway.
 *  3. `visualBgMayBeDark`: true when the slide has background images and the CSS
 *     background-color fell back to white.  In that case the actual visual
 *     background is provided by an image (possibly dark), so compositing over
 *     white is inaccurate.  Suppress when the composited result is "light"
 *     (all channels > 200) because applying a near-white opaque highlight on a
 *     dark visual background looks wrong / hides text.
 *  4. Also suppress when both the highlight and the text color are light (>200) —
 *     additional safety net for the image-backed-dark case when textColor is known.
 */
function computeHighlight(
  backgroundColor: string | undefined,
  textColor: string | undefined,
  slideBg: string,
  visualBgMayBeDark = false,
): string | undefined {
  if (!backgroundColor) return undefined
  const composited = compositeOver(backgroundColor, slideBg)
  const hex = rgbToHex(composited)
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  // Parse slide bg for contrast check
  const bghex = rgbToHex(slideBg)
  const br = parseInt(bghex.slice(0, 2), 16)
  const bgG = parseInt(bghex.slice(2, 4), 16)
  const bgB = parseInt(bghex.slice(4, 6), 16)
  // Suppress: composited color is too close to bg — highlight would be invisible
  if (Math.max(Math.abs(r - br), Math.abs(g - bgG), Math.abs(b - bgB)) < 10)
    return undefined
  // Suppress: image-backed dark slide (css bg is white but visual bg may be dark).
  // A light opaque highlight on a dark visual bg looks wrong.
  if (visualBgMayBeDark && r > 200 && g > 200 && b > 200) return undefined
  // Fallback: suppress when both highlight and text are light — safety net for
  // image-backed dark slides where code text color may not be pure white.
  if (r > 200 && g > 200 && b > 200 && textColor) {
    const tc = rgbToHex(textColor)
    const tr = parseInt(tc.slice(0, 2), 16)
    const tg = parseInt(tc.slice(2, 4), 16)
    const tb = parseInt(tc.slice(4, 6), 16)
    if (tr > 200 && tg > 200 && tb > 200) return undefined
  }
  return hex
}

export function placeElement(
  slide: PptxGenJS.Slide,
  el: SlideElement,
  slideW = 0,
  slideH = 0,
  slideBg = 'rgb(255, 255, 255)',
  visualBgMayBeDark = false,
): void {
  const x = pxToInches(el.x)
  const y = pxToInches(el.y)
  const w = pxToInches(el.width)
  const rawH = pxToInches(el.height)
  // Clamp height for text elements so they never extend beyond the slide area.
  const h =
    slideH > 0 && TEXT_ELEMENT_TYPES.has(el.type)
      ? Math.min(rawH, Math.max(0.01, pxToInches(slideH) - y))
      : rawH

  // Shorthand that carries slideBg into every toTextProps call.
  const toTP = (r: TextRun) => toTextProps(r, slideBg, visualBgMayBeDark)

  switch (el.type) {
    case 'heading': {
      // Draw border-left bar FIRST (z-order: behind text)
      const headingBorderW =
        el.borderLeft && el.borderLeft.width > 0
          ? pxToInches(el.borderLeft.width)
          : 0
      if (headingBorderW > 0) {
        slide.addShape('rect', {
          x,
          y,
          w: headingBorderW,
          h,
          fill: { color: rgbToHex(el.borderLeft!.color) },
          line: { color: rgbToHex(el.borderLeft!.color) },
        })
      }
      // For full-width headings (spanning most of the slide), extend the text
      // box to the slide boundary.  Font metric differences between Chrome and
      // PowerPoint (e.g. DirectWrite vs Skia) can make the same text measure
      // slightly wider in PPTX, causing a single-line heading to wrap to two
      // lines.  Extending to the maximum available width absorbs this variance.
      // The heuristic: heading left edge ≤ 15 % of slide width AND right edge
      // ≥ 85 % of slide width → use slide_width − x_offset − 16 px buffer.
      const isFullWidthHeading =
        slideW > 0 && el.x < slideW * 0.15 && el.x + el.width > slideW * 0.85
      const headingTextW = isFullWidthHeading
        ? Math.max(0.01, pxToInches(slideW - el.x - 16) - headingBorderW)
        : Math.max(0.01, w - headingBorderW)
      // Draw text shifted right so it doesn't overlap the border-left bar.
      // Apply CSS padding as text-box inset (same as blockquote).
      const headingInset = computeTextInset(el.style)
      slide.addText(
        el.runs.map(toTP),
        {
          x: x + headingBorderW,
          y,
          w: headingTextW,
          h,
          margin: headingInset,
          valign: 'top',
          align: el.style.textAlign as PptxGenJS.HAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style),
        },
      )
      // Draw border-bottom as a thin rectangle directly below the heading.
      // Solid borders are rendered as a filled rect; dashed/dotted borders use
      // the same fill-omit / dashType pattern as container borderBottom.
      if (el.borderBottom && el.borderBottom.width > 0) {
        const bh = pxToInches(el.borderBottom.width)
        const bbColor = rgbToHex(el.borderBottom.color)
        const bbDash = cssBorderStyleToDash(el.borderBottom.style)
        slide.addShape('rect', {
          x,
          y: y + h,
          w,
          h: bh,
          ...(bbDash
            ? {
                line: {
                  color: bbColor,
                  width: Math.max(0.25, pxToPoints(el.borderBottom.width)),
                  dashType: bbDash,
                },
              }
            : {
                fill: { color: bbColor },
                line: { color: bbColor, width: 0.25 },
              }),
        })
      }
      break
    }

    case 'paragraph': {
      // Absorb PowerPoint font-metric variance for wide paragraphs.
      // DirectWrite (PPTX) metrics can be slightly wider than Chrome's Skia,
      // causing text that fits on one line in HTML to wrap in PPTX.
      // Heuristic: when the paragraph's right edge is beyond 70 % of the slide
      // width AND the paragraph is itself wider than 25 % of the slide width,
      // extend the text box by up to 32 px (capped at the slide boundary).
      // This gives slack for single-word overflow without reshaping multi-line
      // blocks significantly.
      const paraRightEdge = el.x + el.width
      const paraW =
        slideW > 0 &&
        paraRightEdge > slideW * 0.7 &&
        el.width > slideW * 0.25
          ? Math.max(w, pxToInches(Math.min(el.width + 32, slideW - el.x - 8)))
          : w
      slide.addText(
        el.runs.map(toTP),
        {
          x,
          y,
          w: paraW,
          h,
          margin: computeTextInset(el.style),
          valign: el.valign ?? 'top',
          align: el.style.textAlign as PptxGenJS.HAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style),
        },
      )
      break
    }

    case 'header':
    case 'footer':
      slide.addText(
        el.runs.map(toTP),
        {
          x,
          y,
          w,
          h,
          margin: 0,
          valign: 'top',
          align: el.style.textAlign as PptxGenJS.HAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style),
        },
      )
      break

    case 'blockquote':
      if (el.borderLeft && el.borderLeft.width > 0) {
        const bw = pxToInches(el.borderLeft.width)
        slide.addShape('rect', {
          x,
          y,
          w: bw,
          h,
          fill: { color: rgbToHex(el.borderLeft.color) },
        })
        // Apply CSS padding as text-box inset so the text is properly spaced
        // from the border-left bar.  paddingLeft provides the gap between the
        // bar and the text content; top/bottom padding aligns the first line.
        slide.addText(
          el.runs.map(toTP),
          {
            x: x + bw,
            y,
            w: w - bw,
            h,
            margin: computeTextInset(el.style),
            valign: 'top',
            align: el.style.textAlign as PptxGenJS.HAlign,
            lineSpacingMultiple: computeLineSpacing(el.style),
            paraSpaceBefore: 0,
            charSpacing: computeCharSpacing(el.style),
          },
        )
      } else {
        slide.addText(
          el.runs.map(toTP),
          {
            x,
            y,
            w,
            h,
            margin: computeTextInset(el.style),
            valign: 'top',
            align: el.style.textAlign as PptxGenJS.HAlign,
            lineSpacingMultiple: computeLineSpacing(el.style),
            paraSpaceBefore: 0,
            charSpacing: computeCharSpacing(el.style),
          },
        )
      }
      break

    case 'list':
      slide.addText(
        el.items.flatMap((item, index) =>
          toListTextProps(item, el.ordered, index < el.items.length - 1, slideBg, visualBgMayBeDark),
        ),
        {
          x,
          y,
          w,
          h,
          margin: 0,
          valign: 'top',
          align: el.style.textAlign as PptxGenJS.HAlign,
          lineSpacingMultiple: computeLineSpacing(el.style),
          paraSpaceBefore: 0,
          charSpacing: computeCharSpacing(el.style),
        },
      )
      break

    case 'table': {
      // Compute table-level margin from CSS padding of the first cell.
      // When padding data is available (from dom-walker), use it directly;
      // otherwise fall back to the hand-tuned default that works well for
      // Marp's default theme (top/bottom 0.1in, left/right 0.05in).
      //
      // Note: PptxGenJS addTable margin uses CSS order [top, right, bottom, left],
      // NOT the OOXML bodyPr order [left, right, bottom, top] used by addText margin.
      const refCell = el.rows[0]?.cells[0]
      // Backward compatibility: old dom-walker versions did not extract paddingTop.
      // When the field is missing (undefined), fall back to the fixed defaults.
      const hasCellPadding =
        refCell?.style.paddingTop !== undefined &&
        refCell?.style.paddingTop !== null
      const tableFallbackMargin: [number, number, number, number] =
        hasCellPadding
          ? [
              pxToInches(refCell.style.paddingTop!),
              pxToInches(refCell.style.paddingRight ?? 0),
              pxToInches(refCell.style.paddingBottom ?? 0),
              pxToInches(refCell.style.paddingLeft ?? 0),
            ]
          : [0.1, 0.05, 0.1, 0.05]
      slide.addTable(
        el.rows.map((row) =>
          row.cells.map((cell) => {
            // Per-cell margin from CSS padding (overrides table-level margin)
            const cellMargin: [number, number, number, number] | undefined =
              cell.style.paddingTop !== undefined
                ? [
                    pxToInches(cell.style.paddingTop ?? 0),
                    pxToInches(cell.style.paddingRight ?? 0),
                    pxToInches(cell.style.paddingBottom ?? 0),
                    pxToInches(cell.style.paddingLeft ?? 0),
                  ]
                : undefined
            // Use styled runs if available, otherwise plain text
            if (cell.runs && cell.runs.length > 0) {
              const cellOpts: Record<string, any> = {
                align: cell.style.textAlign as PptxGenJS.HAlign,
                ...(cellMargin ? { margin: cellMargin } : {}),
              }
              if (!isTransparent(cell.style.backgroundColor)) {
                cellOpts.fill = { color: rgbToHex(cell.style.backgroundColor) }
              }
              if (
                cell.style.borderColor &&
                !isTransparent(cell.style.borderColor)
              ) {
                cellOpts.border = {
                  pt: 1,
                  color: rgbToHex(cell.style.borderColor),
                }
              }
              // Effective background for inline highlight computation:
              // use the cell fill colour when set, otherwise the slide bg.
              const cellEffBg = !isTransparent(cell.style.backgroundColor)
                ? cell.style.backgroundColor
                : slideBg
              return {
                text: cell.runs.map((r) => {
                  const hl = computeHighlight(
                    r.backgroundColor,
                    r.color,
                    cellEffBg,
                    visualBgMayBeDark,
                  )
                  return {
                    text: sanitizeText(r.text),
                    options: {
                      color: rgbToHex(r.color),
                      fontSize: pxToPoints(r.fontSize ?? cell.style.fontSize),
                      fontFace: cleanFontFamily(
                        r.fontFamily ?? cell.style.fontFamily,
                        r.text,
                      ),
                      bold:
                        r.bold ?? cell.isHeader ?? cell.style.fontWeight >= 600,
                      italic: r.italic,
                      ...(hl ? { highlight: hl } : {}),
                    },
                  }
                }),
                options: cellOpts,
              }
            }
            // Fallback: plain text
            const cellOpts: Record<string, any> = {
              bold: cell.isHeader || cell.style.fontWeight >= 600,
              color: rgbToHex(cell.style.color),
              fontSize: pxToPoints(cell.style.fontSize),
              fontFace: cleanFontFamily(cell.style.fontFamily, cell.text),
              align: cell.style.textAlign as PptxGenJS.HAlign,
              ...(cellMargin ? { margin: cellMargin } : {}),
            }
            if (!isTransparent(cell.style.backgroundColor)) {
              cellOpts.fill = { color: rgbToHex(cell.style.backgroundColor) }
            }
            if (
              cell.style.borderColor &&
              !isTransparent(cell.style.borderColor)
            ) {
              cellOpts.border = {
                pt: 1,
                color: rgbToHex(cell.style.borderColor),
              }
            }
            return { text: sanitizeText(cell.text), options: cellOpts }
          }),
        ),
        {
          x,
          y,
          w,
          autoPage: false,
          // Preserve HTML column proportions when per-column widths are available
          ...(el.colWidths &&
          el.colWidths.length > 0 &&
          el.colWidths.every((cw) => cw > 0)
            ? {
                // Add a proportional per-column slack to absorb PPTX/Chrome
                // font-metric variance.  DirectWrite (PPTX) renders bold text
                // slightly wider than Chrome's Skia.  A fixed absolute slack
                // was insufficient for longer header strings (e.g. "Column 2
                // (center-aligned)") where the absolute pixel variance at the
                // DirectWrite level can exceed 8 px.  Scaling each column to
                // 105% of the browser-measured width covers the observed ~3%
                // variance across all header cell lengths and font sizes while
                // adding a manageable ~5% total table width overhead.
                // Guard: if scaled column total would exceed the table width,
                // scale back proportionally so the table fits within the slide.
                colW: (() => {
                  const scaled = el.colWidths.map((cw) => cw * DIRECTWRITE_COL_WIDTH_FACTOR)
                  const scaledSum = scaled.reduce((a, b) => a + b, 0)
                  // Guard: prevent the scaled columns from extending beyond the
                  // slide canvas right edge.  Compare against available width
                  // (slide width minus the table's left offset) — NOT against
                  // el.width, which would cancel the 1.05x for all full-row
                  // tables where sum(colWidths) ≈ el.width.
                  const available = SLIDE_CANVAS_W_PX - el.x
                  const clamp = scaledSum > available ? available / scaledSum : 1
                  return scaled.map((cw) => pxToInches(cw * clamp))
                })(),
              }
            : {}),
          // Cell margin derived from CSS padding of the first cell.
          // Per-cell margins (set above) override this when available.
          // Fallback matches Marp's default table cell CSS padding.
          margin: tableFallbackMargin,
        },
      )
      break
    }

    case 'code': {
      // Background rectangle for code blocks
      if (!isTransparent(el.style.backgroundColor)) {
        slide.addShape('rect', {
          x,
          y,
          w,
          h,
          fill: { color: rgbToHex(el.style.backgroundColor) },
        })
      }
      // Code blocks: always use el.text (raw textContent) as the source of truth
      // for line structure. Syntax-highlighted el.runs skips whitespace-only text
      // nodes (blank lines between code sections), so blank lines would be lost
      // when runs are used. Plain monospace text preserves all newlines correctly.
      slide.addText(sanitizeText(el.text), {
        x,
        y,
        w,
        h,
        margin: 0,
        fontFace: 'Courier New',
        fontSize: pxToPoints(el.style.fontSize),
        color: rgbToHex(el.style.color),
        valign: 'top',
        paraSpaceBefore: 0,
      })
      break
    }

    case 'image': {
      const imgOpts: PptxGenJS.ImageProps = {
        x,
        y,
        w,
        h,
        ...resolveImageSource(el.src),
      }
      slide.addImage(imgOpts)
      break
    }

    case 'container': {
      const bg = el.style?.backgroundColor
      const borderWidth = el.style?.borderWidth ?? 0
      const borderColor = el.style?.borderColor
      const borderRadius = el.style?.borderRadius ?? 0
      const borderLeft = el.style?.borderLeft
      const hasBoxShadow = el.style?.boxShadow === true
      const hasBackground = !isTransparent(bg)
      const hasBorder =
        borderWidth > 0 && !!borderColor && !isTransparent(borderColor)

      // Map CSS border-style to PptxGenJS dashType
      const borderDashType = cssBorderStyleToDash(el.style?.borderStyle)

      // Determine effective line (border) for the shape.
      // box-shadow → thin grey line to simulate card elevation.
      const lineStyle: Record<string, any> | undefined = hasBorder
        ? {
            color: rgbToHex(borderColor!),
            width: pxToPoints(borderWidth),
            ...(borderDashType ? { dashType: borderDashType } : {}),
          }
        : hasBoxShadow
          ? { color: 'CCCCCC', width: 0.5 }
          : undefined

      if (hasBackground || hasBorder || hasBoxShadow) {
        // Use 'roundRect' shape type when border-radius is set
        const shapeType = borderRadius > 0 ? 'roundRect' : 'rect'
        // rectRadius is 0-1: convert px radius relative to the smaller dimension
        const minDim = Math.min(el.width, el.height)
        const rectRadius =
          borderRadius > 0
            ? Math.min(0.5, borderRadius / (minDim / 2))
            : undefined
        slide.addShape(shapeType as PptxGenJS.ShapeType, {
          x,
          y,
          w,
          h,
          fill: hasBackground ? { color: rgbToHex(bg!) } : { type: 'none' },
          ...(lineStyle ? { line: lineStyle } : {}),
          ...(rectRadius !== undefined ? { rectRadius } : {}),
        })
      }
      // Draw border-left bar (e.g. note-box left accent bar)
      if (borderLeft && borderLeft.width > 0) {
        const bw = pxToInches(borderLeft.width)
        slide.addShape('rect', {
          x,
          y,
          w: bw,
          h,
          fill: { color: rgbToHex(borderLeft.color) },
          line: { color: rgbToHex(borderLeft.color) },
        })
      }
      // Draw border-bottom as a thin line at the element's bottom edge
      // (e.g. row separators, section underlines).
      const borderBottom = el.style?.borderBottom
      if (borderBottom && borderBottom.width > 0) {
        const bbh = pxToInches(borderBottom.width)
        const bbColor = rgbToHex(borderBottom.color)
        // Map CSS border-style to PptxGenJS dashType
        const bbDash = cssBorderStyleToDash(borderBottom.style)
        slide.addShape('rect', {
          x,
          y: y + h,
          w,
          h: bbh,
          // For dashed/dotted borders: omit fill so PptxGenJS generates
          // <a:noFill/>.  Passing fill:{type:'none'} is a truthy object and
          // routes through genXmlColorSelection() which only handles 'solid' —
          // it outputs nothing, leaving the shape with the slide-theme default
          // fill (potentially opaque) which would mask the dash pattern.
          // Omitting fill makes options.fill falsy → PptxGenJS emits <a:noFill/>.
          // For solid borders, a filled rect renders cleanly as a solid rule.
          ...(bbDash
            ? {
                line: {
                  color: bbColor,
                  width: Math.max(0.25, pxToPoints(borderBottom.width)),
                  dashType: bbDash,
                },
              }
            : {
                fill: { color: bbColor },
                line: { color: bbColor, width: 0.25 },
              }),
        })
      }
      // Badge/chip text: render runs centered inside the shape.
      // extractInlineBadgeShapes captures badge text directly so it aligns
      // perfectly with the badge background shape, avoiding the misalignment
      // that occurs when text is placed from the parent paragraph's text flow.
      if (
        el.runs &&
        el.runs.length > 0 &&
        el.runs.some((r) => !r.breakLine && r.text.trim() !== '')
      ) {
        slide.addText(
          el.runs.map(toTP),
          {
            x,
            y,
            w,
            h,
            margin: 0,
            valign: 'middle',
            align: 'center',
            lineSpacingMultiple: 1,
            paraSpaceBefore: 0,
          },
        )
      }
      // Recursively place children.
      // When the container has a visible background, strip redundant highlight
      // from children's text runs whose backgroundColor matches the container
      // fill.  The shape already provides the visual background; keeping the
      // same colour as a text highlight causes visible artefacts (colour bleed
      // on slight positioning mismatches).
      if (hasBackground) {
        const bgHex = rgbToHex(bg!)
        for (const child of el.children ?? []) {
          if ('runs' in child && Array.isArray((child as any).runs)) {
            for (const r of (child as any).runs as TextRun[]) {
              if (
                !r.breakLine &&
                r.backgroundColor &&
                rgbToHex(r.backgroundColor) === bgHex
              ) {
                r.backgroundColor = undefined
              }
            }
          }
        }
      }
      for (const child of el.children ?? []) {
        placeElement(slide, child, slideW, slideH, slideBg, visualBgMayBeDark)
      }
      break
    }
  }
}

export function toTextProps(
  run: TextRun,
  slideBg = 'rgb(255, 255, 255)',
  visualBgMayBeDark = false,
): PptxGenJS.TextProps {
  // Explicit break run (inserted by extractTextRuns for block boundaries / <br>)
  if (run.breakLine) {
    return { text: '', options: { breakLine: true } }
  }

  const text = sanitizeText(run.text)
  const highlight = computeHighlight(
    run.backgroundColor,
    run.color,
    slideBg,
    visualBgMayBeDark,
  )

  return {
    text,
    options: {
      color: rgbToHex(run.color),
      fontSize: pxToPoints(run.fontSize ?? 16),
      fontFace: cleanFontFamily(run.fontFamily, run.text),
      bold: run.bold,
      italic: run.italic,
      underline: run.underline ? { style: 'sng' } : undefined,
      strike: run.strikethrough ? 'sngStrike' : undefined,
      hyperlink: run.hyperlink ? { url: run.hyperlink } : undefined,
      highlight,
    },
  }
}

export function toListTextProps(
  item: ListItem,
  ordered = false,
  breakAfter = false,
  slideBg = 'rgb(255, 255, 255)',
  visualBgMayBeDark = false,
): PptxGenJS.TextProps[] {
  const bulletOption = createListBulletOption(item, ordered)

  if (item.runs.length === 0) {
    return [
      {
        text: sanitizeText(item.text) || ' ',
        options: {
          bullet: bulletOption,
          indentLevel: item.level,
          breakLine: breakAfter,
        },
      },
    ]
  }

  // Split runs at <br> boundaries so each continuation line becomes its own
  // paragraph with the correct left margin.
  //
  // Background: PptxGenJS's breakLine:true creates a new <a:p> (not <a:br/>).
  // When opts.align is set (always the case for list addText calls), a truthy
  // bullet option does NOT trigger a paragraph boundary — only breakLine does.
  //
  // Strategy: end each non-last group with breakLine:true so the next group
  // starts in an empty arrTexts.  For continuation groups use
  // bullet:{char:'\u200B'} (zero-width space — invisible) so PptxGenJS emits a
  // bullet paragraph with the correct marL, matching the text-start position of
  // the first bullet paragraph (PowerPoint Shift+Enter / soft-return behaviour).
  const groups: TextRun[][] = [[]]
  for (const run of item.runs) {
    if (run.breakLine) {
      groups.push([])
    } else {
      groups[groups.length - 1].push(run)
    }
  }

  const result: PptxGenJS.TextProps[] = []
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g]
    if (group.length === 0) continue
    const isContinuation = g > 0
    const isLastGroup = g === groups.length - 1
    const groupBullet = isContinuation
      ? createListBulletOption(item, ordered, true)
      : bulletOption
    for (let r = 0; r < group.length; r++) {
      const run = group[r]
      const isLastRun = r === group.length - 1
      // End this paragraph when:
      //   - last run of a non-last group  →  clears arrTexts for the next group
      //   - last run of the last group AND breakAfter  →  inter-item separator
      const needsBreakLine = isLastRun && (!isLastGroup || breakAfter)
      result.push({
        text: sanitizeText(run.text),
        options: {
          // Always propagate bullet and indentLevel to every run in the group.
          // PptxGenJS v4.x emits <a:pPr> for each TextProp in the same
          // paragraph. LibreOffice uses the *last* <a:pPr>, so without
          // propagation the last run's pPr resets the bullet with <a:buNone/>.
          // PowerPoint uses the *first* <a:pPr> and is unaffected by this change.
          bullet: groupBullet,
          indentLevel: item.level,
          ...(needsBreakLine ? { breakLine: true } : {}),
          color: rgbToHex(run.color),
          fontSize: pxToPoints(run.fontSize ?? 16),
          fontFace: cleanFontFamily(run.fontFamily, run.text),
          bold: run.bold,
          italic: run.italic,
          highlight: computeHighlight(
            run.backgroundColor,
            run.color,
            slideBg,
            visualBgMayBeDark,
          ),
        },
      })
    }
  }
  return result
}
