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
 *
 * PowerPoint adds internal leading (~15% of font size) on top of
 * lineSpacingMultiple, so a raw CSS ratio of 1.5 renders wider than CSS
 * line-height:1.5.  We divide by the internal leading factor (1.15) to
 * compensate, clamping at 0.85 so text never collapses.
 */
function computeLineSpacing(
  style: TextStyle,
  multiLine?: boolean,
): number | undefined {
  const { lineHeight, fontSize } = style
  if (!lineHeight || !fontSize || lineHeight <= 0 || fontSize <= 0)
    return undefined
  const m = lineHeight / fontSize
  if (m < 0.5 || m > 4) return undefined
  // Compensate for PowerPoint internal leading.
  // Single-line elements: divide by 1.15 (conservative, avoids baseline shift).
  // Multi-line elements: divide by 1.20 (tighter line spacing to match CSS
  // rendering where accumulated inter-line gaps make PPTX text appear shifted).
  const divisor = multiLine ? 1.2 : 1.15
  const adjusted = Math.max(0.85, m / divisor)
  return Math.round(adjusted * 100) / 100
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

/**
 * Map CSS list-style-type to PptxGenJS bullet number style.
 * Returns undefined for unknown types (falls back to arabicPeriod).
 */
function cssListStyleToNumberStyle(
  cssType: string | undefined,
): string | undefined {
  if (!cssType) return undefined
  switch (cssType) {
    case 'decimal':
      return 'arabicPeriod'
    case 'lower-alpha':
    case 'lower-latin':
      return 'alphaLcPeriod'
    case 'upper-alpha':
    case 'upper-latin':
      return 'alphaUcPeriod'
    case 'lower-roman':
      return 'romanLcPeriod'
    case 'upper-roman':
      return 'romanUcPeriod'
    default:
      return undefined
  }
}

/**
 * Map CSS list-style-type to a bullet character code for unordered lists.
 * Returns undefined for default disc bullet (PptxGenJS default).
 */
function cssListStyleToBulletChar(
  cssType: string | undefined,
): string | undefined {
  if (!cssType) return undefined
  switch (cssType) {
    case 'circle':
      return '25E6' // ◦ (white bullet — smaller than ○ U+25CB)
    case 'square':
      return '25AA' // ▪ (black small square)
    case 'none':
      return '200B' // zero-width space (invisible)
    default:
      return undefined // disc = PptxGenJS default bullet
  }
}

function createListBulletOption(
  item: ListItem,
  ordered: boolean,
  continuation = false,
  startNumber?: number,
  listStyleType?: string,
): boolean | Record<string, any> {
  const extraIndent = item.leadingOffset ? pxToPoints(item.leadingOffset) : 0
  const indent =
    extraIndent > 0
      ? Math.round((DEFAULT_BULLET_INDENT_PT + extraIndent) * 100) / 100
      : undefined

  // Resolve effective list style from item or parent list
  const effectiveStyle = item.listStyleType ?? listStyleType

  if (continuation) {
    return {
      characterCode: '200B',
      ...(indent !== undefined ? { indent } : {}),
    }
  }

  // Determine effective ordered/unordered for this specific item.
  // A nested <ul> inside an <ol> will have a bullet-type listStyleType
  // (disc/circle/square/none) — treat as unordered regardless of parent.
  // Conversely a nested <ol> inside <ul> will have a number-type style.
  const isUnorderedStyle =
    effectiveStyle === 'disc' ||
    effectiveStyle === 'circle' ||
    effectiveStyle === 'square' ||
    effectiveStyle === 'none'
  const isOrderedStyle =
    effectiveStyle === 'decimal' ||
    effectiveStyle === 'lower-alpha' ||
    effectiveStyle === 'lower-latin' ||
    effectiveStyle === 'upper-alpha' ||
    effectiveStyle === 'upper-latin' ||
    effectiveStyle === 'lower-roman' ||
    effectiveStyle === 'upper-roman'
  const effectiveOrdered = isUnorderedStyle
    ? false
    : isOrderedStyle
      ? true
      : ordered

  if (effectiveOrdered) {
    const numberStyle =
      cssListStyleToNumberStyle(effectiveStyle) ?? 'arabicPeriod'
    return {
      type: 'number',
      style: numberStyle,
      ...(startNumber !== undefined ? { numberStartAt: startNumber } : {}),
      ...(indent !== undefined ? { indent } : {}),
    }
  }

  const bulletChar = cssListStyleToBulletChar(effectiveStyle)
  if (bulletChar) {
    return {
      characterCode: bulletChar,
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
      !firstBg.backgroundSizeContain &&
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
    //
    // Structural fidelity policy (see Design Principles in instructions):
    //
    // HTML-specified structural elements — including inline <code> background
    // highlights — are ALWAYS rendered, even when the result is visually
    // imperfect.  For example, on a split-background slide where text sits over
    // a colored image, the code highlight will appear as a near-white solid box
    // against the image.  This is accepted: the structure says "code", so PPTX
    // says "code highlight".  Per-element suppression based on visual
    // approximation violates the "browser is source of truth" principle.
    //
    // The only permitted suppression is for full-slide dark-background slides:
    // when a non-contain ![bg] image covers ≥ 80 % of slide width and CSS
    // bg-color is white, the slide's visual background is the image, not white,
    // and the near-white composited highlight would be effectively invisible.
    // `backgroundSizeContain` images (![bg fit]) are excluded: the image is
    // letterboxed in the centre; margins remain on the CSS background (white).
    const cssIsFallbackWhite =
      !slideData.background ||
      rgbToHex(slideData.background).toUpperCase() === 'FFFFFF'
    const visualBgMayBeDark =
      bgImages.some(
        (bg) =>
          bg.url !== '' &&
          !bg.fromCssFallback &&
          !bg.backgroundSizeContain &&
          bg.width >= slideData.width * 0.8,
      ) && cssIsFallbackWhite

    const slideBgColor = slideData.background ?? 'rgb(255, 255, 255)'

    // Build the spatial container-text association map.  Text elements whose
    // bounding boxes are contained within a visible card container are embedded
    // directly into the shape and must not be rendered as standalone text boxes.
    const containerAssoc = associateContainerText(
      slideData.elements,
      slideData.width,
      slideData.height,
    )
    const embeddedByContainer = new Set(
      Array.from(containerAssoc.values()).flat(),
    )

    // Filter out text elements that would be mostly clipped by overflow:hidden.
    // HTML sections clip at slideHeight; PPTX text boxes don't clip.
    // Only text elements are filtered — containers/images may intentionally bleed.
    const visibleElements = slideData.elements.filter(
      (el) =>
        !TEXT_ELEMENT_TYPES.has(el.type) || slideData.height - el.y >= 20,
    )

    const groups = groupAdjacentTextElements(visibleElements)
    for (const group of groups) {
      // Skip elements that will be rendered as part of a container shape.
      const active = group.filter((el) => !embeddedByContainer.has(el))
      if (active.length === 0) continue

      if (active.length === 1) {
        placeElement(
          slide,
          active[0],
          slideData.width,
          slideData.height,
          slideBgColor,
          visualBgMayBeDark,
          containerAssoc,
        )
      } else {
        placeGroupedTextElements(
          slide,
          active,
          slideData.width,
          slideData.height,
          slideBgColor,
          visualBgMayBeDark,
        )
      }
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

// Element types eligible for text grouping (adjacent elements with similar
// position and width merged into a single text box).
// Headings, tables, code blocks, images, containers, and header/footer are excluded
// because they have unique rendering logic or should remain visually separate.
const GROUPABLE_TYPES = new Set<string>([
  'paragraph',
  'blockquote',
])

/**
 * Child element types that can be embedded as text inside a container shape.
 * Nested containers, images, tables, and code blocks are excluded because
 * they have their own rendering paths or positional requirements.
 */
const EMBEDDABLE_IN_SHAPE = new Set([
  'paragraph',
  'heading',
  'list',
  'blockquote',
])

/**
 * Returns true when all children of the container are plain text elements that
 * can be embedded directly into the shape text body.
 *
 * Containers are excluded when they:
 * - have no children or already carry badge/chip text (`runs`)
 * - have border-left / border-bottom decorations that require separate shapes
 * - contain non-text children (images, nested containers, tables, etc.)
 * - have children laid out horizontally (flex row), since embedding would
 *   collapse distinct x-positions into a single text flow with wrong margins
 */
function isSimpleTextContainer(
  el: import('./types').ContainerElement,
): boolean {
  if (!el.children || el.children.length === 0) return false
  if (el.runs && el.runs.length > 0) return false
  if (el.style.borderBottom) return false
  if (!el.children.every((c) => EMBEDDABLE_IN_SHAPE.has(c.type))) return false

  // Detect horizontal flex layout: if children span a wide x-range they must
  // be placed as individual text boxes, not merged into one text flow.
  if (el.children.length > 1) {
    const xs = el.children.map((c) => c.x)
    const xSpan = Math.max(...xs) - Math.min(...xs)
    if (xSpan > 50) return false
  }

  return true
}

/**
 * Compute the [left, right, bottom, top] text inset (in inches) for embedding
 * children inside a container shape, derived from the spatial offset between
 * the container's bounding box and its first/last child.
 *
 * Returns 0 when the inset is negligible or cannot be determined.
 */
function computeContainerInset(
  el: import('./types').ContainerElement,
): [number, number, number, number] | 0 {
  const { children } = el
  if (!children || children.length === 0) return 0

  const firstChild = children[0]
  const lastChild = children[children.length - 1]

  // Offset of first child from container origin.
  const topPx = Math.max(0, firstChild.y - el.y)
  const leftPx = Math.max(0, firstChild.x - el.x)
  const rightPx = Math.max(
    0,
    el.x + el.width - (firstChild.x + firstChild.width),
  )
  const bottomPx = Math.max(
    0,
    el.y + el.height - (lastChild.y + lastChild.height),
  )

  const top = topPx * 0.75 // px → pt (1px = 0.75pt, same scale as computeTextInset)
  const left = leftPx * 0.75
  const right = rightPx * 0.75
  const bottom = bottomPx * 0.75

  // PptxGenJS margin order: [left, right, bottom, top]
  return top || left || right || bottom ? [left, right, bottom, top] : 0
}

/**
 * Collect all text runs from a container's text children into a flat array
 * suitable for `slide.addText()`.  Children are separated by a breakLine run.
 */
function buildContainerEmbeddedRuns(
  el: import('./types').ContainerElement,
  slideBg: string,
  visualBgMayBeDark: boolean,
): PptxGenJS.TextProps[] {
  const toTP = (r: TextRun) => toTextProps(r, slideBg, visualBgMayBeDark)
  const allRuns: PptxGenJS.TextProps[] = []

  for (let i = 0; i < el.children.length; i++) {
    const child = el.children[i]
    if (i > 0) allRuns.push({ text: '', options: { breakLine: true } })

    if (child.type === 'list') {
      const listEl = child as import('./types').ListElement
      let topLevelCount = 0
      const listRuns = listEl.items.flatMap((item, idx) => {
        const startNum =
          item.level === 0 && listEl.ordered
            ? (listEl.startNumber ?? 1) + topLevelCount
            : undefined
        if (item.level === 0) topLevelCount++
        return toListTextProps(
          item,
          listEl.ordered,
          idx < listEl.items.length - 1,
          slideBg,
          visualBgMayBeDark,
          startNum,
          listEl.listStyleType,
        )
      })
      allRuns.push(...listRuns)
    } else if ('runs' in child && Array.isArray((child as any).runs)) {
      allRuns.push(...(child as any).runs.map(toTP))
    }
  }
  return allRuns
}

/** Maximum vertical gap (px) between two elements to allow grouping. */
const MAX_GROUP_GAP_PX = 64

// ── Spatial container-text association ───────────────────────────────
//
// In the DOM walker output, visible "card" container elements typically have
// empty children[] — the text that visually appears inside a card is extracted
// at the slide top level and positioned via CSS.  To couple them into a single
// PPTX shape+text object we use spatial containment: a text element whose
// bounding box fits within a container's bounding box is "owned" by that
// container.

/** Tolerance (px) for containment check: minor font/layout drift is allowed. */
const CONTAINMENT_SLOP_PX = 4

/** Returns true when `text` is spatially contained within `container`. */
function isContainedIn(
  text: SlideElement,
  container: SlideElement,
): boolean {
  // Strict containment with small tolerance
  const strict =
    text.x >= container.x - CONTAINMENT_SLOP_PX &&
    text.y >= container.y - CONTAINMENT_SLOP_PX &&
    text.x + text.width <=
      container.x + container.width + CONTAINMENT_SLOP_PX &&
    text.y + text.height <=
      container.y + container.height + CONTAINMENT_SLOP_PX
  if (strict) return true

  // Near-colocated: text origin is very close to container origin and sizes
  // are similar.  This catches code-block patterns where Marp wraps a <pre>
  // inside a styled <div> and the inner text bbox slightly exceeds the outer
  // container due to font-metric / overflow:visible differences.
  const dx = Math.abs(text.x - container.x)
  const dy = Math.abs(text.y - container.y)
  const dw = Math.abs(text.width - container.width)
  const dh = Math.abs(text.height - container.height)
  if (dx <= 16 && dy <= 16 && dw <= 16 && dh <= 16) return true

  return false
}

/**
 * Returns true when the container is a candidate for text embedding:
 * - Has a visible appearance (background, border, or box-shadow)
 * - No borderLeft / borderBottom decorations (those need separate shapes)
 * - No badge/chip text (`runs` already carries rendered text)
 * - Not a full-slide background element (>90% of both dimensions)
 */
function isEmbeddableContainer(
  el: import('./types').ContainerElement,
  slideW: number,
  slideH: number,
): boolean {
  if (el.runs && el.runs.length > 0) return false
  // borderBottom needs a separate shape (horizontal rule) — exclude.
  // borderLeft is allowed: text is embedded with extra margin-left; the bar
  // is drawn separately but the text+background become a single object.
  if (el.style?.borderBottom) return false
  // Containers with complex children (non-embeddable types like nested containers,
  // images, tables, code blocks) must NOT be targets for spatial text association.
  // Using the embedded path would skip recursive placement of those children,
  // causing content loss (e.g. 2-column grids inside chat bubbles).
  if (el.children && el.children.length > 0 && !isSimpleTextContainer(el)) {
    return false
  }
  const hasBackground = !isTransparent(el.style?.backgroundColor)
  const hasBorderLeft = !!(
    el.style?.borderLeft &&
    el.style.borderLeft.width > 0
  )
  const hasBorder =
    (el.style?.borderWidth ?? 0) > 0 &&
    !!el.style?.borderColor &&
    !isTransparent(el.style.borderColor!)
  const hasShadow = el.style?.boxShadow === true
  if (!hasBackground && !hasBorder && !hasShadow && !hasBorderLeft) return false
  if (slideW > 0 && slideH > 0) {
    if (el.width >= slideW * 0.9 && el.height >= slideH * 0.9) return false
  }
  return true
}

/**
 * Build a mapping from each embeddable container to the slide-level text
 * elements spatially contained inside it.
 *
 * Each text element is assigned to the SMALLEST enclosing container so that
 * nested cards assign text to the inner card, not the outer wrapper.
 */
export function associateContainerText(
  elements: SlideElement[],
  slideW: number,
  slideH: number,
): Map<import('./types').ContainerElement, SlideElement[]> {
  const result = new Map<
    import('./types').ContainerElement,
    SlideElement[]
  >()

  const containers = elements.filter(
    (el): el is import('./types').ContainerElement =>
      el.type === 'container' &&
      isEmbeddableContainer(
        el as import('./types').ContainerElement,
        slideW,
        slideH,
      ),
  )
  if (containers.length === 0) return result

  // Smallest container first so inner cards take priority over outer wrappers
  const sorted = [...containers].sort(
    (a, b) => a.width * a.height - b.width * b.height,
  )

  const textElements = elements.filter((el) =>
    EMBEDDABLE_IN_SHAPE.has(el.type),
  )
  const assigned = new Set<SlideElement>()

  for (const container of sorted) {
    const inside = textElements.filter(
      (el) => !assigned.has(el) && isContainedIn(el, container),
    )
    if (inside.length > 0) {
      inside.sort((a, b) => a.y - b.y) // natural reading order
      result.set(container, inside)
      inside.forEach((el) => assigned.add(el))
    }
  }
  return result
}

/**
 * Collect text runs from an ordered array of slide-level text elements into a
 * flat `TextProps[]` suitable for `slide.addText()`.
 * Adjacent elements are separated by a breakLine run.
 * Runs whose `backgroundColor` matches `containerBg` are stripped (the shape
 * already provides the fill; keeping the highlight causes colour bleed).
 */
function buildRunsFromElements(
  elements: SlideElement[],
  containerBg: string | undefined,
  slideBg: string,
  visualBgMayBeDark: boolean,
): PptxGenJS.TextProps[] {
  const toTP = (r: TextRun) => toTextProps(r, slideBg, visualBgMayBeDark)
  const bgHex = containerBg ? rgbToHex(containerBg) : undefined
  const allRuns: PptxGenJS.TextProps[] = []

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]
    if (i > 0) allRuns.push({ text: '', options: { breakLine: true } })

    if (el.type === 'list') {
      const listEl = el as import('./types').ListElement
      let topLevelCount = 0
      const listRuns = listEl.items.flatMap((item, idx) => {
        const startNum =
          item.level === 0 && listEl.ordered
            ? (listEl.startNumber ?? 1) + topLevelCount
            : undefined
        if (item.level === 0) topLevelCount++
        return toListTextProps(
          item,
          listEl.ordered,
          idx < listEl.items.length - 1,
          slideBg,
          visualBgMayBeDark,
          startNum,
          listEl.listStyleType,
        )
      })
      allRuns.push(...listRuns)
    } else if ('runs' in el && Array.isArray((el as any).runs)) {
      const runs: TextRun[] = (el as any).runs as TextRun[]
      const mapped = runs.map((r) => {
        if (bgHex && !r.breakLine && r.backgroundColor && rgbToHex(r.backgroundColor) === bgHex) {
          return { ...r, backgroundColor: undefined }
        }
        return r
      })
      allRuns.push(...mapped.map(toTP))
    }
  }
  return allRuns
}

/**
 * Compute text inset ([left, right, bottom, top] in pt) for a container whose
 * text elements are the given slide-level elements.
 *
 * Returns 0 when the inset is negligible.
 */
function computeInsetFromElements(
  container: import('./types').ContainerElement,
  elements: SlideElement[],
): [number, number, number, number] | 0 {
  if (elements.length === 0) return 0
  const firstEl = elements[0]
  const lastEl = elements[elements.length - 1]

  const topPx = Math.max(0, firstEl.y - container.y)
  const leftPx = Math.max(
    0,
    Math.min(...elements.map((e) => e.x - container.x)),
  )
  const rightPx = Math.max(
    0,
    Math.min(
      ...elements.map(
        (e) => container.x + container.width - (e.x + e.width),
      ),
    ),
  )
  const bottomPx = Math.max(
    0,
    container.y + container.height - (lastEl.y + lastEl.height),
  )

  // 1px = 0.75pt (same scale used in computeTextInset)
  const top = topPx * 0.75
  const left = leftPx * 0.75
  const right = rightPx * 0.75
  const bottom = bottomPx * 0.75

  // PptxGenJS margin order: [left, right, bottom, top]
  return top || left || right || bottom ? [left, right, bottom, top] : 0
}
/** Maximum horizontal position difference (px) to consider elements aligned. */
const MAX_X_DIFF_PX = 5
/** Maximum width difference (px) to consider elements similarly sized. */
const MAX_WIDTH_DIFF_PX = 5

/**
 * Returns true when two adjacent text elements can be merged into a single
 * text box.
 */
function canMergeElements(a: SlideElement, b: SlideElement): boolean {
  if (!GROUPABLE_TYPES.has(a.type) || !GROUPABLE_TYPES.has(b.type)) return false
  // Don't merge multi-line paragraphs — cross-engine text wrapping
  // differences (e.g. CJK character metrics) can change the rendered
  // height of the paragraph, propagating Y-position errors to all
  // subsequent paragraphs in the group.
  if (a.type === 'paragraph' && 'style' in a) {
    const s = (a as any).style as import('./types').TextStyle
    if (s.lineHeight > 0 && a.height > s.lineHeight * 1.5) return false
  }
  if (b.type === 'paragraph' && 'style' in b) {
    const s = (b as any).style as import('./types').TextStyle
    if (s.lineHeight > 0 && b.height > s.lineHeight * 1.5) return false
  }
  // Elements must be roughly the same column (x-aligned, similar width)
  if (Math.abs(a.x - b.x) > MAX_X_DIFF_PX) return false
  if (Math.abs(a.width - b.width) > MAX_WIDTH_DIFF_PX) return false
  // Vertical gap must be small
  const gap = b.y - (a.y + a.height)
  if (gap < 0 || gap > MAX_GROUP_GAP_PX) return false
  // Skip elements with decorations that need dedicated rendering
  if (a.type === 'heading' && ('borderBottom' in a || 'borderLeft' in a)) {
    const h = a as import('./types').HeadingElement
    if ((h.borderBottom && h.borderBottom.width > 0) || (h.borderLeft && h.borderLeft.width > 0)) return false
  }
  if (b.type === 'heading' && ('borderBottom' in b || 'borderLeft' in b)) {
    const h = b as import('./types').HeadingElement
    if ((h.borderBottom && h.borderBottom.width > 0) || (h.borderLeft && h.borderLeft.width > 0)) return false
  }
  if (a.type === 'blockquote' && 'borderLeft' in a) {
    const bq = a as import('./types').BlockquoteElement
    if (bq.borderLeft && bq.borderLeft.width > 0) return false
  }
  if (b.type === 'blockquote' && 'borderLeft' in b) {
    const bq = b as import('./types').BlockquoteElement
    if (bq.borderLeft && bq.borderLeft.width > 0) return false
  }
  return true
}

/**
 * Group adjacent text elements that can be merged into a single text box.
 * Non-groupable elements are returned as single-element arrays.
 */
export function groupAdjacentTextElements(
  elements: SlideElement[],
): SlideElement[][] {
  if (elements.length === 0) return []
  const groups: SlideElement[][] = [[elements[0]]]
  for (let i = 1; i < elements.length; i++) {
    const current = elements[i]
    const lastGroup = groups[groups.length - 1]
    const lastEl = lastGroup[lastGroup.length - 1]
    if (canMergeElements(lastEl, current)) {
      lastGroup.push(current)
    } else {
      groups.push([current])
    }
  }
  return groups
}

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
  containerAssoc?: Map<import('./types').ContainerElement, SlideElement[]>,
): void {
  // Skip text elements mostly clipped by overflow:hidden on the <section>.
  // Containers/images are exempt (may bleed intentionally).
  if (slideH > 0 && TEXT_ELEMENT_TYPES.has(el.type) && slideH - el.y < 20) return

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

  // Detect multi-line elements for tighter line spacing compensation.
  const elLineHeight =
    'style' in el && 'lineHeight' in el.style ? el.style.lineHeight : 0
  const isMultiLine = elLineHeight > 0 && el.height > elLineHeight * 1.5

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
      // Apply DirectWrite width compensation to ALL headings.  Full-width
      // headings extend to the slide boundary; others get a 5% buffer capped
      // at the available space (same factor as DIRECTWRITE_COL_WIDTH_FACTOR).
      const headingTextW = isFullWidthHeading
        ? Math.max(0.01, pxToInches(slideW - el.x - 16) - headingBorderW)
        : slideW > 0
          ? Math.max(
              0.01,
              pxToInches(
                Math.min(
                  el.width * DIRECTWRITE_COL_WIDTH_FACTOR,
                  slideW - el.x,
                ),
              ) - headingBorderW,
            )
          : Math.max(0.01, w - headingBorderW)
      // Draw text shifted right so it doesn't overlap the border-left bar.
      // Apply CSS padding as text-box inset (same as blockquote).
      const headingInset = computeTextInset(el.style)
      // Compensate for CSS half-leading: shift text box up by half-leading
      // so the first baseline aligns with the CSS-rendered position.
      const headingHalfLeading =
        el.style.lineHeight > 0 && el.style.fontSize > 0
          ? (el.style.lineHeight - el.style.fontSize) / 2
          : 0
      const headingY = headingHalfLeading > 0 ? pxToInches(el.y - headingHalfLeading) : y
      slide.addText(
        el.runs.map(toTP),
        {
          x: x + headingBorderW,
          y: headingY,
          w: headingTextW,
          h,
          margin: headingInset,
          valign: 'top',
          align: el.style.textAlign as PptxGenJS.HAlign,
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
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
      // Absorb PowerPoint font-metric variance for paragraphs.
      // DirectWrite (PPTX) metrics can be ~3-5% wider than Chrome's Skia,
      // causing text that fits on one line in HTML to wrap in PPTX.
      // Apply width buffer only to SINGLE-LINE paragraphs; multi-line
      // paragraphs keep the exact HTML width so line break positions match.
      const noWrap = el.style.whiteSpace === 'nowrap'
      const isSingleLine =
        el.style.lineHeight > 0 && el.height <= el.style.lineHeight * 1.5
      let paraW: number
      if (noWrap) {
        // For nowrap paragraphs, extend the text box to the slide boundary.
        // wrap:false prevents line breaks regardless, but a wider box
        // ensures all text is visible and selectable for editing.
        paraW = slideW > 0
          ? Math.max(w, pxToInches(Math.min(el.width * 1.15, slideW - el.x)))
          : w
      } else if (slideW > 0 && isSingleLine) {
        // Single-line: extend width to prevent unwanted wrapping in PPTX.
        const extendedW = el.width * DIRECTWRITE_COL_WIDTH_FACTOR
        const maxW = slideW - el.x - 4
        paraW = Math.max(w, pxToInches(Math.min(extendedW, maxW)))
      } else {
        paraW = w
      }
      // Compensate for CSS half-leading: shift text box up so the first
      // baseline aligns with the CSS-rendered position.
      const paraHalfLeading =
        el.style.lineHeight > 0 && el.style.fontSize > 0
          ? (el.style.lineHeight - el.style.fontSize) / 2
          : 0
      const paraY = paraHalfLeading > 0 ? pxToInches(el.y - paraHalfLeading) : y
      slide.addText(
        el.runs.map(toTP),
        {
          x,
          y: paraY,
          w: paraW,
          h,
          margin: computeTextInset(el.style),
          valign: el.valign ?? 'top',
          align: el.style.textAlign as PptxGenJS.HAlign,
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: computeCharSpacing(el.style),
          ...(noWrap ? { wrap: false } : {}),
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
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
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
            lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
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
            lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
            charSpacing: computeCharSpacing(el.style),
          },
        )
      }
      break

    case 'list': {
      // Compensate for CSS half-leading: shift text box up so the first
      // bullet aligns with the CSS-rendered position.
      const halfLeadingPx =
        el.style.lineHeight > 0 && el.style.fontSize > 0
          ? (el.style.lineHeight - el.style.fontSize) / 2
          : 0
      const yShiftPx = halfLeadingPx * 1.0
      const listY = pxToInches(el.y - yShiftPx)
      const listH = Math.max(0.01, h + pxToInches(yShiftPx))

      let topLevelCount = 0
      slide.addText(
        el.items.flatMap((item, index) => {
          // Compute per-item startNumber so every top-level item carries an
          // explicit numberStartAt.  PptxGenJS does NOT auto-increment from the
          // first item's startAt — each run without numberStartAt resets to 1.
          // Nested items (level > 0) belong to a sub-list and always start at 1.
          const startNum =
            item.level === 0 && el.ordered
              ? (el.startNumber ?? 1) + topLevelCount
              : undefined
          if (item.level === 0) topLevelCount++
          return toListTextProps(
            item,
            el.ordered,
            index < el.items.length - 1,
            slideBg,
            visualBgMayBeDark,
            startNum,
            el.listStyleType,
          )
        }),
        {
          x,
          y: listY,
          w,
          h: listH,
          margin: 0,
          valign: 'top',
          align: el.style.textAlign as PptxGenJS.HAlign,
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: computeCharSpacing(el.style),
        },
      )
      break
    }

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
                ...(cell.colspan && cell.colspan > 1 ? { colspan: cell.colspan } : {}),
                ...(cell.rowspan && cell.rowspan > 1 ? { rowspan: cell.rowspan } : {}),
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

              // Helper: convert a single TextRun into a PptxGenJS text prop
              const runToTextProp = (r: TextRun) => {
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
              }

              // Prefer structured paragraphs when available; fall back to
              // flat runs with breakLine markers for backward compatibility.
              let textArray: any[]
              if (cell.paragraphs && cell.paragraphs.length > 0) {
                textArray = cell.paragraphs.flatMap((para, pIdx) => {
                  const paraRuns: any[] = para.runs
                    .filter((r) => !r.breakLine)
                    .map(runToTextProp)
                  // Emit at least one entry per paragraph so PptxGenJS creates
                  // an empty <a:p> (visible as a blank line in the cell).
                  if (paraRuns.length === 0) {
                    paraRuns.push({ text: ' ', options: {
                      fontSize: pxToPoints(cell.style.fontSize),
                    } })
                  }
                  // Add breakLine to last run of non-final paragraphs
                  if (pIdx < cell.paragraphs!.length - 1) {
                    const last = paraRuns[paraRuns.length - 1]
                    paraRuns[paraRuns.length - 1] = {
                      ...last,
                      options: { ...last.options, breakLine: true },
                    }
                  }
                  return paraRuns
                })
              } else {
                textArray = cell.runs.map((r) => {
                  if (r.breakLine) {
                    return { text: '', options: { breakLine: true } }
                  }
                  return runToTextProp(r)
                })
              }

              return {
                text: textArray,
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
              ...(cell.colspan && cell.colspan > 1 ? { colspan: cell.colspan } : {}),
              ...(cell.rowspan && cell.rowspan > 1 ? { rowspan: cell.rowspan } : {}),
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
      // Code blocks: single object with shape background + text inside.
      // This keeps background and text as one selectable/movable unit in PPTX.
      const hasCodeBg = !isTransparent(el.style.backgroundColor)
      const codeShapeOpts: Record<string, any> = {
        x,
        y,
        w,
        h,
        margin: 0,
        valign: 'top',
        paraSpaceBefore: 0,
        paraSpaceAfter: 0,
        autoFit: false,
        wrap: false,
        ...(hasCodeBg
          ? {
              shape: 'rect' as PptxGenJS.ShapeType,
              fill: { color: rgbToHex(el.style.backgroundColor) },
            }
          : {}),
      }

      // Compute font-size cap: ensure all lines fit within the shape height.
      // This guards against auto-scaling mismatches where the extracted
      // fontSize is too large for the PPTX shape box.
      const LINE_SPACING = 1.2
      const codeNumLines = el.runs
        ? el.runs.filter((r) => r.breakLine).length + 1
        : (el.text?.split('\n').length ?? 1)
      const shapeHeightPt = h * 72 // h is in inches
      const baseFontSizePt = pxToPoints(el.style.fontSize)
      const maxFontSizePt =
        codeNumLines > 1
          ? shapeHeightPt / (codeNumLines * LINE_SPACING)
          : shapeHeightPt
      const codeFontScale =
        baseFontSizePt > maxFontSizePt ? maxFontSizePt / baseFontSizePt : 1

      // Code blocks: prefer syntax-highlighted runs when available.
      // Each run carries its own colour from the highlight.js / Prism theme,
      // and newlines are represented as breakLine runs so blank lines are
      // preserved.  Font uses cleanFontFamily to pick the best monospace
      // (e.g. Consolas on Windows) from the CSS font stack.
      const codeFontFace = cleanFontFamily(
        el.style.fontFamily,
        el.text,
      )
      if (el.runs && el.runs.length > 0) {
        slide.addText(
          el.runs.map((r) => {
            if (r.breakLine) {
              return { text: '', options: { breakLine: true } }
            }
            return {
              text: sanitizeText(r.text),
              options: {
                color: rgbToHex(r.color),
                fontSize:
                  pxToPoints(r.fontSize ?? el.style.fontSize) * codeFontScale,
                fontFace: codeFontFace,
                bold: r.bold,
                italic: r.italic,
              },
            }
          }),
          codeShapeOpts,
        )
      } else {
        // Fallback: plain monospace text (no syntax highlighting)
        slide.addText(sanitizeText(el.text), {
          ...codeShapeOpts,
          fontFace: codeFontFace,
          fontSize: pxToPoints(el.style.fontSize) * codeFontScale,
          color: rgbToHex(el.style.color),
        })
      }
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

      // Pre-compute shape type and radius for both paths.
      const shapeType: PptxGenJS.ShapeType =
        (borderRadius > 0 ? 'roundRect' : 'rect') as PptxGenJS.ShapeType
      const minDim = Math.min(el.width, el.height)
      const rectRadius =
        borderRadius > 0
          ? Math.min(0.5, borderRadius / (minDim / 2))
          : undefined

      const isVisibleShape = hasBackground || hasBorder || hasBoxShadow

      // ── Embedded text path ────────────────────────────────────────
      // Primary: spatial association — text elements at slide level whose
      // bounding boxes are contained within this container (most common case).
      // Secondary: children-based — text elements in el.children[] (rare).
      // Both paths produce a single addText() call with the shape as background,
      // coupling text and shape into one PPTX object.
      const spatialText = containerAssoc?.get(
        el as import('./types').ContainerElement,
      )
      const useEmbeddedPath =
        isVisibleShape &&
        ((spatialText && spatialText.length > 0) ||
          isSimpleTextContainer(el as import('./types').ContainerElement))

      if (useEmbeddedPath) {
        let runs: PptxGenJS.TextProps[]
        let margin: [number, number, number, number] | 0
        let firstStyle: import('./types').TextStyle | undefined

        if (spatialText && spatialText.length > 0) {
          // ── Spatial path ──
          runs = buildRunsFromElements(
            spatialText,
            hasBackground ? bg : undefined,
            slideBg,
            visualBgMayBeDark,
          )
          margin = computeInsetFromElements(
            el as import('./types').ContainerElement,
            spatialText,
          )
          const firstEl = spatialText[0]
          firstStyle =
            'style' in firstEl
              ? ((firstEl as any).style as import('./types').TextStyle)
              : undefined
        } else {
          // ── Children path (fallback) ──
          if (hasBackground) {
            const bgHex = rgbToHex(bg!)
            for (const child of el.children) {
              if ('runs' in child && Array.isArray((child as any).runs)) {
                for (const r of (child as any).runs as TextRun[]) {
                  if (!r.breakLine && r.backgroundColor && rgbToHex(r.backgroundColor) === bgHex) {
                    r.backgroundColor = undefined
                  }
                }
              }
            }
          }
          runs = buildContainerEmbeddedRuns(el, slideBg, visualBgMayBeDark)
          margin = computeContainerInset(el)
          const firstChild = el.children[0]
          firstStyle =
            'style' in firstChild
              ? ((firstChild as any).style as import('./types').TextStyle)
              : undefined
        }

        // Add extra margin-left for border-left bar so text doesn't overlap it.
        const borderLeft = el.style?.borderLeft
        if (borderLeft && borderLeft.width > 0) {
          const extraLeft = borderLeft.width * 0.75 + 2 // pt: bar width + gap
          if (Array.isArray(margin)) {
            margin = [margin[0] + extraLeft, margin[1], margin[2], margin[3]]
          } else {
            margin = [extraLeft, 0, 0, 0]
          }
        }

        // For small badge-like containers (circle/pill), centre text.
        const isBadgeLike =
          el.width <= 80 &&
          el.height <= 80 &&
          borderRadius >= 50

        slide.addText(runs, {
          shape: shapeType,
          x,
          y,
          w,
          h,
          fill: hasBackground ? { color: rgbToHex(bg!) } : { type: 'none' },
          ...(lineStyle ? { line: lineStyle } : {}),
          ...(rectRadius !== undefined ? { rectRadius } : {}),
          margin,
          valign: isBadgeLike ? 'middle' : 'top',
          align: isBadgeLike
            ? 'center'
            : (firstStyle?.textAlign as PptxGenJS.HAlign),
          lineSpacingMultiple: isBadgeLike
            ? 1
            : firstStyle
              ? computeLineSpacing(firstStyle)
              : undefined,
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: firstStyle ? computeCharSpacing(firstStyle) : undefined,
          autoFit: false,
          wrap: true,
        })

        // Draw border-left bar as a separate thin rect (decorative).
        if (borderLeft && borderLeft.width > 0) {
          const bw = pxToInches(borderLeft.width)
          const blColor = rgbToHex(borderLeft.color)
          slide.addShape('rect', {
            x,
            y,
            w: bw,
            h,
            fill: { color: blColor },
            line: { color: blColor, width: 0.25 },
          })
        }
        break
      }

      // ── Badge/chip path: single object (shape + centred text) ──────
      // Containers with runs carry inline badge/chip text directly.  Emit as
      // a single addText with shape option to produce one editable PPTX object.
      const hasBadgeRuns =
        el.runs &&
        el.runs.length > 0 &&
        el.runs.some((r) => !r.breakLine && r.text.trim() !== '')
      if (hasBadgeRuns && isVisibleShape) {
        slide.addText(
          el.runs!.map(toTP),
          {
            shape: shapeType,
            x,
            y,
            w,
            h,
            fill: hasBackground ? { color: rgbToHex(bg!) } : { type: 'none' },
            ...(lineStyle ? { line: lineStyle } : {}),
            ...(rectRadius !== undefined ? { rectRadius } : {}),
            margin: 0,
            valign: 'middle',
            align: 'center',
            lineSpacingMultiple: 1,
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
          },
        )
        break
      }

      // ── Standard path (fallback) ──────────────────────────────────
      if (isVisibleShape) {
        slide.addShape(shapeType, {
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
      // Apply container-text association at child level (same logic as slide
      // level) so that badge circles / card backgrounds among siblings absorb
      // their adjacent text into a single PPTX object.
      const childElements = el.children ?? []
      const childAssoc = associateContainerText(
        childElements,
        slideW,
        slideH,
      )
      const childEmbedded = new Set(
        Array.from(childAssoc.values()).flat(),
      )
      for (const child of childElements) {
        if (childEmbedded.has(child)) continue
        placeElement(
          slide,
          child,
          slideW,
          slideH,
          slideBg,
          visualBgMayBeDark,
          childAssoc,
        )
      }
      break
    }
  }
}

/**
 * Place a group of adjacent text elements as a single text box.
 * The bounding box is the union of all elements in the group.
 * Each element's runs become a paragraph (separated by breakLine) within
 * the unified text box, preserving the visual layout.
 */
function placeGroupedTextElements(
  slide: PptxGenJS.Slide,
  group: SlideElement[],
  slideW: number,
  slideH: number,
  slideBg: string,
  visualBgMayBeDark: boolean,
): void {
  // Compute union bounding box
  const minX = Math.min(...group.map((e) => e.x))
  const minY = Math.min(...group.map((e) => e.y))
  const maxX = Math.max(...group.map((e) => e.x + e.width))
  const maxY = Math.max(...group.map((e) => e.y + e.height))
  const x = pxToInches(minX)
  // Compensate for CSS half-leading using first element's style
  const firstStyle =
    'style' in group[0] ? (group[0] as any).style as import('./types').TextStyle : undefined
  const groupHalfLeading =
    firstStyle && firstStyle.lineHeight > 0 && firstStyle.fontSize > 0
      ? (firstStyle.lineHeight - firstStyle.fontSize) / 2
      : 0
  const y = pxToInches(minY - groupHalfLeading)
  const w = pxToInches(maxX - minX)
  const h =
    slideH > 0
      ? Math.min(pxToInches(maxY - minY), Math.max(0.01, pxToInches(slideH) - y))
      : pxToInches(maxY - minY)

  const toTP = (r: TextRun) => toTextProps(r, slideBg, visualBgMayBeDark)

  // Build runs array: each element's runs become a paragraph, separated by
  // a breakLine run between elements.
  const allRuns: PptxGenJS.TextProps[] = []
  for (let i = 0; i < group.length; i++) {
    const el = group[i]
    if (i > 0) {
      // Paragraph separator between elements
      allRuns.push({ text: '', options: { breakLine: true } })
    }
    if (el.type === 'list') {
      const listEl = el as import('./types').ListElement
      let topLevelCount = 0
      const listRuns = listEl.items.flatMap((item, idx) => {
        const startNum =
          item.level === 0 && listEl.ordered
            ? (listEl.startNumber ?? 1) + topLevelCount
            : undefined
        if (item.level === 0) topLevelCount++
        return toListTextProps(
          item,
          listEl.ordered,
          idx < listEl.items.length - 1,
          slideBg,
          visualBgMayBeDark,
          startNum,
          listEl.listStyleType,
        )
      })
      allRuns.push(...listRuns)
    } else if ('runs' in el && Array.isArray((el as any).runs)) {
      allRuns.push(...(el as any).runs.map(toTP))
    }
  }

  // Use the first element's style for text box options
  const first = group[0]
  const style =
    'style' in first ? (first as any).style as import('./types').TextStyle : undefined
  slide.addText(allRuns, {
    x,
    y,
    w,
    h,
    margin: style ? computeTextInset(style) : 0,
    valign: 'top',
    align: style?.textAlign as PptxGenJS.HAlign,
    lineSpacingMultiple: style ? computeLineSpacing(style) : undefined,
    paraSpaceBefore: 0,
    paraSpaceAfter: 0,
    charSpacing: style ? computeCharSpacing(style) : undefined,
  })
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
      subscript: run.subscript || undefined,
      superscript: run.superscript || undefined,
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
  startNumber?: number,
  listStyleType?: string,
): PptxGenJS.TextProps[] {
  const bulletOption = createListBulletOption(item, ordered, false, startNumber, listStyleType)

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
      ? createListBulletOption(item, ordered, true, undefined, listStyleType)
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
          subscript: run.subscript || undefined,
          superscript: run.superscript || undefined,
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
