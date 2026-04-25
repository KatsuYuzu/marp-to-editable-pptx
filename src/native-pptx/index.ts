import { readFile, writeFile, access } from 'node:fs/promises'
import { pathToFileURL, fileURLToPath } from 'node:url'
import puppeteer, { type Browser, type Page } from 'puppeteer-core'
import { DOM_WALKER_SCRIPT } from './dom-walker-script.generated'
import { buildPptx } from './slide-builder'
import type { ImageElement, SlideData, SlideElement } from './types'

export interface NativePptxOptions {
  /** Absolute path to the HTML file rendered by Marp CLI. */
  htmlPath: string
  /** Absolute path to a Chromium-based browser executable. */
  browserPath: string
  /** Slide viewport width in pixels (default: 1280). */
  width?: number
  /** Slide viewport height in pixels (default: 720). */
  height?: number
  /** If set, dump extracted SlideData[] JSON to this path for diagnostics. */
  debugJsonPath?: string
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function shouldHidePaginationPseudoText(
  rawContent: string | null | undefined,
  paginationValue: string | null | undefined,
  paginationTotal?: string | null | undefined,
): boolean {
  const normalizedPaginationValue = paginationValue?.trim()
  if (!normalizedPaginationValue) return false

  const normalizedContent = rawContent?.trim()
  if (
    !normalizedContent ||
    normalizedContent === 'none' ||
    normalizedContent === 'normal'
  ) {
    return false
  }

  const strippedContent = normalizedContent.replace(/^['"]|['"]$/g, '').trim()
  if (!strippedContent) return false

  const escapedPaginationValue = escapeRegExp(normalizedPaginationValue)
  const normalizedPaginationTotal = paginationTotal?.trim()
  const exactValuePattern = new RegExp(`^\\(?0*${escapedPaginationValue}\\)?$`)
  if (exactValuePattern.test(strippedContent)) return true

  const labeledValuePattern = new RegExp(
    `^(?:page|slide|p\\.?|#)\\s*0*${escapedPaginationValue}$`,
    'i',
  )
  if (labeledValuePattern.test(strippedContent)) return true

  if (!normalizedPaginationTotal) return false

  const escapedPaginationTotal = escapeRegExp(normalizedPaginationTotal)
  const pagedFractionPattern = new RegExp(
    `^(?:(?:page|slide|p\\.?|#)\\s*)?0*${escapedPaginationValue}\\s*(?:/|of)\\s*0*${escapedPaginationTotal}$`,
    'i',
  )
  return pagedFractionPattern.test(strippedContent)
}

/**
 * Generate an editable PPTX buffer from a Marp-rendered HTML file.
 *
 * 1. Launch a headless browser via puppeteer-core
 * 2. Load the HTML and wait for rendering
 * 3. Inject the DOM walker script and extract structured slide data
 * 4. Build a PPTX presentation from the extracted data
 * 5. Return the PPTX as a Node.js Buffer
 */
export async function generateNativePptx(
  opts: NativePptxOptions,
): Promise<Buffer> {
  const { htmlPath, browserPath, width = 1280, height = 720 } = opts

  let browser: Browser | undefined

  try {
    browser = await puppeteer.launch({
      executablePath: browserPath,
      headless: true,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
      ],
    })

    const page = await browser.newPage()
    await page.setViewport({ width, height })

    // Navigate to the HTML file using a file:// URL so that relative paths
    // (e.g. images referenced as "./image.png") resolve correctly against
    // the directory containing the HTML file.
    const fileUrl = pathToFileURL(htmlPath).href
    await page.goto(fileUrl, { waitUntil: 'networkidle0' })

    // After networkidle0, external scripts (e.g. mermaid.js from CDN) have
    // finished downloading and started executing.  However, script-based
    // renderers like mermaid use async Promises/microtasks to insert SVGs into
    // the DOM.  Give them time to complete before we walk the DOM.
    await new Promise((r) => setTimeout(r, 1000))

    // Hide bespoke presentation UI elements so they don't appear in
    // Puppeteer screenshots used for CSS-filtered backgrounds.
    // The OSC overlay sits on top of slides; note panels are off-slide
    // but could affect page layout if not hidden.
    await page.addStyleTag({
      content:
        '.bespoke-marp-osc,[data-bespoke-marp-osc],.bespoke-marp-note{display:none!important}',
    })

    // Inject the DOM walker as a self-contained IIFE script, then call
    // extractSlides() in the browser context.
    //
    // We cannot use page.evaluate(fn) with a direct function reference because
    // webpack's esbuild minimizer (keepNames: true) injects module-scope
    // helpers like `t(fn, name)` into function bodies.  After toString()
    // serialization those references are lost and cause ReferenceError in the
    // browser.  Instead, the DOM walker is compiled separately by esbuild into
    // a standalone IIFE (see scripts/generate-dom-walker-script.js) and
    // embedded as a string constant that is safe to inject via addScriptTag.
    await page.addScriptTag({ content: DOM_WALKER_SCRIPT })
    const slides: SlideData[] = await page.evaluate(() =>
      (globalThis as any).extractSlides(),
    )

    // After structured extraction, hide only the pagination text itself so
    // rasterized backgrounds do not duplicate the native PPTX slide number
    // while decorative bars / pills / borders on the same pseudo-element stay
    // visible.
    if (slides.some((slide) => slide.sourceHasPagination)) {
      const paginationMatches: Array<{
        index: number
        paginationValue: string
        paginationTotal: string
        afterContent: string
      }> = await page.evaluate(() =>
        Array.from(
          document.querySelectorAll('section[data-marpit-pagination]'),
        ).map((section, index) => ({
          index,
          paginationValue:
            section.getAttribute('data-marpit-pagination')?.trim() ?? '',
          paginationTotal:
            section.getAttribute('data-marpit-pagination-total')?.trim() ?? '',
          afterContent: getComputedStyle(section, '::after').content ?? '',
        })),
      )

      const matchedPaginationIndexes = paginationMatches
        .filter(({ afterContent, paginationValue, paginationTotal }) =>
          shouldHidePaginationPseudoText(
            afterContent,
            paginationValue,
            paginationTotal,
          ),
        )
        .map(({ index }) => index)

      if (matchedPaginationIndexes.length > 0) {
        await page.evaluate((matchedIndexes: number[]) => {
          const sections = Array.from(
            document.querySelectorAll('section[data-marpit-pagination]'),
          )
          for (const index of matchedIndexes) {
            sections[index]?.setAttribute(
              'data-native-pptx-hide-pagination-after',
              'true',
            )
          }
        }, matchedPaginationIndexes)
        await page.addStyleTag({
          content:
            'section[data-native-pptx-hide-pagination-after="true"]::after{color:transparent!important;-webkit-text-fill-color:transparent!important;text-shadow:none!important}',
        })
      }
    }

    // Handle missing local image files before any rasterization pass:
    // - Content images (ImageElement): screenshot the browser's own broken-image
    //   rendering so the missing asset is clearly indicated in the PPTX, matching
    //   the visual shown in the HTML/PDF output.
    // - Background images (BgImageData): CSS backgrounds silently disappear when
    //   the URL is missing (no broken-image indicator), so we simply remove those
    //   entries; the slide falls back to its background-color fill.
    const missingUrls = await findMissingLocalUrls(slides)
    await rasterizeSlideTargets(page, buildBrokenContentImageJobs(slides, missingUrls))
    pruneMissingBackgrounds(slides, missingUrls)

    // Rasterize CSS-filtered background images via Puppeteer screenshot.
    // This captures grayscale, brightness, sepia, blur etc. that PptxGenJS
    // cannot reproduce natively.
    await rasterizeSlideTargets(page, buildFilteredBgJobs(slides))
    await rasterizeSlideTargets(page, buildCssFallbackBgJobs(slides))
    await rasterizeSlideTargets(page, buildFilteredContentImageJobs(slides))
    // Rasterize images flagged for screenshot-based capture (e.g. Mermaid SVGs
    // that use <foreignObject> internally and cannot be embedded as-is).
    await rasterizeSlideTargets(page, buildRasterizeImageJobs(slides))
    // Rasterize partial-width background images (e.g. ![bg right:30%]).
    // CSS background-size:cover crops differently than PPTX stretch-to-fill,
    // so we screenshot the rendered figure region for accurate reproduction.
    await rasterizeSlideTargets(page, buildPartialBgJobs(slides))

    // Diagnostic dump: save extracted data as JSON for comparing HTML → JSON → PPTX
    if (opts.debugJsonPath) {
      await writeFile(
        opts.debugJsonPath,
        JSON.stringify(slides, null, 2),
        'utf-8',
      )
    }

    // Resolve remaining local image paths to data: URLs before building the PPTX.
    // Missing files were already handled above (content images screenshotted,
    // background images removed), so any failure here is an unexpected I/O error
    // and will propagate as an error dialog.
    await resolveImageUrls(slides)

    // Build PPTX from extracted data
    const pptx = buildPptx(slides)
    const output = await pptx.write({ outputType: 'nodebuffer' })

    return Buffer.from(output as ArrayBuffer)
  } finally {
    await browser?.close()
  }
}

// ---------------------------------------------------------------------------
// Unified rasterization engine
// ---------------------------------------------------------------------------

/** Timing for page hash-navigation to settle before screenshots. */
const NAVIGATION_SETTLE_MS = 300
/** Timing after all rasterization before returning control. */
const POST_RASTERIZE_SETTLE_MS = 100

interface RasterizeTarget {
  clip: { x: number; y: number; width: number; height: number }
  /**
   * When true, the clip coordinates are slide-relative (x/y measured from
   * the slide section's top-left corner).  rasterizeSlideTargets queries
   * the current section's getBoundingClientRect() AFTER navigating to that
   * slide and adds the origin offset, so bespoke-mode CSS transforms (which
   * move inactive slides off-screen at extraction time) do not corrupt the
   * clip rectangle.
   */
  slideRelative?: boolean
  /** Store the rasterized base64 data-URL into the originating element. */
  onCapture(dataUrl: string): void
}

interface SlideRasterizeJob {
  slideIdx: number
  targets: RasterizeTarget[]
  /** Prepare page visibility before screenshots (e.g. hide overlapping layers). */
  setup?(page: Page, slideIdx: number): Promise<void>
  /** Restore page visibility after screenshots. */
  teardown?(page: Page, slideIdx: number): Promise<void>
}

/**
 * Navigate to each slide, optionally manipulate visibility, then screenshot
 * every target clip region and store the base64 result via onCapture.
 */
async function rasterizeSlideTargets(
  page: Page,
  jobs: SlideRasterizeJob[],
): Promise<void> {
  if (jobs.length === 0) return

  for (const { slideIdx, targets, setup, teardown } of jobs) {
    await page.evaluate((n: number) => {
      window.location.hash = '#' + n
    }, slideIdx + 1)
    await new Promise<void>((r) => setTimeout(r, NAVIGATION_SETTLE_MS))

    // If any target uses slide-relative coordinates, resolve the target
    // slide's absolute page position.
    //
    // Strategy: first look up the section by its numeric id (e.g. id="12").
    // In bespoke.js HTML, hash navigation has already moved that section to
    // the viewport, so rect.top ≈ 0.  In static Marp HTML (no bespoke.js),
    // hash navigation has no visual effect; the section stays at its absolute
    // page position (e.g. y = slideIdx * slideHeight).  Using
    // rect.top + scrollY converts viewport-relative coords to page-absolute
    // coords, which is what page.screenshot({ clip }) expects.
    //
    // Fallback: most-visible-section heuristic for bespoke.js layouts that
    // do not use numeric section ids.
    let slideOriginX = 0
    let slideOriginY = 0
    if (targets.some((t) => t.slideRelative)) {
      const origin = await page.evaluate((n: number) => {
        const target = document.getElementById(String(n))
        if (target) {
          const r = target.getBoundingClientRect()
          return {
            x: r.left + window.scrollX,
            y: r.top + window.scrollY,
          }
        }
        // Fallback: most visible section (bespoke.js without numeric ids)
        const visibleSections = Array.from(
          document.querySelectorAll<HTMLElement>('section'),
        )
          .map((section) => {
            const rect = section.getBoundingClientRect()
            const visibleWidth =
              Math.min(rect.right, window.innerWidth) - Math.max(rect.left, 0)
            const visibleHeight =
              Math.min(rect.bottom, window.innerHeight) - Math.max(rect.top, 0)
            const visibleArea =
              Math.max(0, visibleWidth) * Math.max(0, visibleHeight)
            return { rect, visibleArea }
          })
          .filter((entry) => entry.visibleArea > 0)

        if (visibleSections.length === 0) {
            console.warn(
              `[rasterize] slide ${n}: section id not found and no visible ` +
                'sections in viewport; using origin (0,0) which may produce ' +
                'an incorrect clip for bespoke.js HTML layouts.',
            )
            return { x: 0, y: 0 }
          }

        visibleSections.sort((left, right) => right.visibleArea - left.visibleArea)
        return {
          x: visibleSections[0].rect.left + window.scrollX,
          y: visibleSections[0].rect.top + window.scrollY,
        }
      }, slideIdx + 1)
      slideOriginX = origin.x
      slideOriginY = origin.y
    }

    try {
      if (setup) await setup(page, slideIdx)
      for (const { clip, slideRelative, onCapture } of targets) {
        const effectiveClip = slideRelative
          ? {
              x: Math.round(slideOriginX + clip.x),
              y: Math.round(slideOriginY + clip.y),
              width: clip.width,
              height: clip.height,
            }
          : clip
        if (effectiveClip.width <= 0 || effectiveClip.height <= 0) continue
        try {
          const raw = await page.screenshot({
            type: 'png',
            clip: effectiveClip,
          })
          onCapture(
            'data:image/png;base64,' + Buffer.from(raw).toString('base64'),
          )
        } catch {
          /* skip — element may be off-screen */
        }
      }
    } finally {
      if (teardown) await teardown(page, slideIdx)
    }
  }

  await page.evaluate(() => {
    window.location.hash = '#1'
  })
  await new Promise<void>((r) => setTimeout(r, POST_RASTERIZE_SETTLE_MS))
}

// ---------------------------------------------------------------------------
// Visibility helpers for rasterization setup/teardown
// ---------------------------------------------------------------------------

const ADVANCED_LAYERS_SELECTOR =
  'section[data-marpit-advanced-background="content"], section[data-marpit-advanced-background="pseudo"]'

async function hideAdvancedLayers(page: Page): Promise<void> {
  await page.evaluate((sel) => {
    document
      .querySelectorAll(sel)
      .forEach((el) =>
        (el as HTMLElement).style.setProperty(
          'visibility',
          'hidden',
          'important',
        ),
      )
  }, ADVANCED_LAYERS_SELECTOR)
}

async function restoreAdvancedLayers(page: Page): Promise<void> {
  await page.evaluate((sel) => {
    document
      .querySelectorAll(sel)
      .forEach((el) => (el as HTMLElement).style.removeProperty('visibility'))
  }, ADVANCED_LAYERS_SELECTOR)
}

async function hideSectionChildren(
  page: Page,
  slideIdx: number,
): Promise<void> {
  // Target the section by its numeric id so this works in both bespoke.js
  // HTML (section in viewport after hash nav) and static Marp HTML (section
  // may be off-screen; viewport-based detection would target the wrong slide).
  // Hide children and check for section existence in a single page.evaluate
  // round-trip to avoid the extra Puppeteer IPC overhead of two sequential calls.
  const found = await page.evaluate((id: string) => {
    const section = document.getElementById(id)
    if (!section) return false
    Array.from(section.children).forEach((el) =>
      (el as HTMLElement).style.setProperty('visibility', 'hidden', 'important'),
    )
    return true
  }, String(slideIdx + 1))
  if (!found) {
    console.warn(
      `[rasterize] hideSectionChildren: section id="${slideIdx + 1}" not found; ` +
        'children not hidden — background screenshot may include slide content.',
    )
  }
}

async function restoreSectionChildren(
  page: Page,
  slideIdx: number,
): Promise<void> {
  // Mirror of hideSectionChildren; see that function for rationale.
  await page.evaluate((id: string) => {
    const section = document.getElementById(id)
    if (!section) return
    Array.from(section.children).forEach((el) =>
      (el as HTMLElement).style.removeProperty('visibility'),
    )
  }, String(slideIdx + 1))
}

// ---------------------------------------------------------------------------
// Image URL resolver and missing-image helpers
// ---------------------------------------------------------------------------

/** Returns true for local file URLs (file:) and absolute paths. */
function isLocalImagePath(url: string): boolean {
  if (!url) return false
  if (url.startsWith('data:')) return false
  if (url.startsWith('http://') || url.startsWith('https://')) return false
  return true
}

/**
 * Convert a local file URL or absolute file path to a base64 data URL.
 * Throws on I/O error — callers are responsible for pre-filtering missing
 * files via findMissingLocalUrls before calling this.
 */
async function fileUrlToDataUrl(url: string): Promise<string> {
  const filePath = url.startsWith('file:') ? fileURLToPath(url) : url
  const buf = await readFile(filePath)
  // Derive a rough MIME type from the extension; default to image/png.
  const ext = filePath.split('.').pop()?.toLowerCase() ?? ''
  const mime =
    ext === 'jpg' || ext === 'jpeg'
      ? 'image/jpeg'
      : ext === 'gif'
        ? 'image/gif'
        : ext === 'webp'
          ? 'image/webp'
          : ext === 'svg'
            ? 'image/svg+xml'
            : 'image/png'
  return `data:${mime};base64,${buf.toString('base64')}`
}

/**
 * Walk all SlideData[] and replace every local-file image URL (file: or
 * absolute path) with a base64 data URL so that PptxGenJS never attempts
 * a synchronous fs.readFileSync.
 *
 * data: URLs (already resolved) and http/https URLs are left untouched.
 * Missing files must be handled before calling this (via findMissingLocalUrls +
 * buildBrokenContentImageJobs + pruneMissingBackgrounds).
 */
async function resolveImageUrls(slides: SlideData[]): Promise<void> {
  const jobs: Promise<void>[] = []

  function resolveEl(el: SlideElement): void {
    if (el.type === 'image' && isLocalImagePath(el.src)) {
      jobs.push(
        fileUrlToDataUrl(el.src).then((d) => {
          el.src = d
        }),
      )
    }
    if ('children' in el && Array.isArray((el as any).children)) {
      for (const child of (el as any).children) resolveEl(child)
    }
  }

  for (const slide of slides) {
    // Background images
    for (const bg of slide.backgroundImages ?? []) {
      if (isLocalImagePath(bg.url)) {
        jobs.push(
          fileUrlToDataUrl(bg.url).then((d) => {
            bg.url = d
          }),
        )
      }
    }
    // Content elements
    for (const el of slide.elements ?? []) resolveEl(el)
  }

  await Promise.all(jobs)
}

/**
 * Collect all local-path image URLs across slides and return a Set of those
 * that cannot be accessed (missing files).  Used to route missing assets to
 * the correct handler before any rasterization pass runs.
 */
async function findMissingLocalUrls(slides: SlideData[]): Promise<Set<string>> {
  const urlsToCheck = new Set<string>()

  function collectImageUrls(elements: SlideElement[]): void {
    for (const el of elements) {
      if (el.type === 'image' && isLocalImagePath(el.src))
        urlsToCheck.add(el.src)
      if ('children' in el && Array.isArray((el as any).children)) {
        collectImageUrls((el as any).children)
      }
    }
  }

  for (const slide of slides) {
    for (const bg of slide.backgroundImages ?? []) {
      if (isLocalImagePath(bg.url)) urlsToCheck.add(bg.url)
    }
    collectImageUrls(slide.elements ?? [])
  }

  const missing = new Set<string>()
  await Promise.all(
    [...urlsToCheck].map(async (url) => {
      try {
        const filePath = url.startsWith('file:') ? fileURLToPath(url) : url
        await access(filePath)
      } catch {
        missing.add(url)
      }
    }),
  )
  return missing
}

/**
 * Build rasterization jobs that screenshot the browser's own broken-image
 * rendering for content images whose source file is missing.
 *
 * Chromium renders a broken-image indicator (icon + alt text / filename) within
 * the element's layout bounds — exactly the same visual shown in HTML and PDF
 * output.  Capturing that screenshot makes the missing asset clearly visible in
 * the exported PPTX rather than producing a confusing solid-color placeholder.
 */
function buildBrokenContentImageJobs(
  slides: SlideData[],
  missingUrls: Set<string>,
): SlideRasterizeJob[] {
  function collectBrokenImages(elements: SlideElement[]): ImageElement[] {
    const result: ImageElement[] = []
    for (const el of elements) {
      if (el.type === 'image' && missingUrls.has(el.src)) result.push(el)
      if ('children' in el && Array.isArray((el as any).children)) {
        result.push(...collectBrokenImages((el as any).children))
      }
    }
    return result
  }

  return slides.flatMap((s, i): SlideRasterizeJob[] => {
    const imgs = collectBrokenImages(s.elements ?? [])
    if (imgs.length === 0) return []
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img): RasterizeTarget => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height),
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl
            },
          }),
        ),
      },
    ]
  })
}

/**
 * Remove background image entries whose source file is missing.
 *
 * CSS `background-image` with an inaccessible URL simply renders nothing
 * (no broken indicator).  Removing the entry reproduces that behaviour:
 * the slide falls back to its solid background-color fill.
 */
function pruneMissingBackgrounds(
  slides: SlideData[],
  missingUrls: Set<string>,
): void {
  for (const slide of slides) {
    slide.backgroundImages = (slide.backgroundImages ?? []).filter(
      (bg) => !missingUrls.has(bg.url),
    )
  }
}

// ---------------------------------------------------------------------------
// Job builders: translate SlideData[] into SlideRasterizeJob[]
// ---------------------------------------------------------------------------

function buildFilteredBgJobs(slides: SlideData[]): SlideRasterizeJob[] {
  return slides.flatMap((s, i): SlideRasterizeJob[] => {
    const bgs = (s.backgroundImages ?? []).filter((b) => b.cssFilter)
    if (bgs.length === 0) return []
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg): RasterizeTarget => ({
            clip: {
              x: Math.round(bg.x),
              y: Math.round(bg.y),
              width: Math.round(bg.width),
              height: Math.round(bg.height),
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl
              delete bg.cssFilter
            },
          }),
        ),
        setup: (p) => hideAdvancedLayers(p),
        teardown: (p) => restoreAdvancedLayers(p),
      },
    ]
  })
}

function buildCssFallbackBgJobs(slides: SlideData[]): SlideRasterizeJob[] {
  return slides.flatMap((s, i): SlideRasterizeJob[] => {
    const bgs = (s.backgroundImages ?? []).filter((b) => b.fromCssFallback)
    if (bgs.length === 0) return []
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg): RasterizeTarget => ({
            clip: {
              x: 0,
              y: 0,
              width: Math.round(s.width),
              height: Math.round(s.height),
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl
              // Keep fromCssFallback=true so buildPptx can tell this was a CSS
              // gradient rather than a user-specified dark bg image (![bg])
              // and can avoid suppressing light inline-code highlights.
            },
          }),
        ),
        setup: (p, idx) => hideSectionChildren(p, idx),
        teardown: (p, idx) => restoreSectionChildren(p, idx),
      },
    ]
  })
}

function collectFilteredContentImages(
  elements: SlideElement[],
): ImageElement[] {
  const result: ImageElement[] = []
  for (const el of elements ?? []) {
    if (el.type === 'image' && el.cssFilter) result.push(el)
    if ('children' in el && Array.isArray((el as any).children)) {
      result.push(...collectFilteredContentImages((el as any).children))
    }
  }
  return result
}

function buildFilteredContentImageJobs(
  slides: SlideData[],
): SlideRasterizeJob[] {
  return slides.flatMap((s, i): SlideRasterizeJob[] => {
    const imgs = collectFilteredContentImages(s.elements)
    if (imgs.length === 0) return []
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img): RasterizeTarget => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height),
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl
              delete img.cssFilter
            },
          }),
        ),
      },
    ]
  })
}

function collectRasterizeImages(elements: SlideElement[]): ImageElement[] {
  const result: ImageElement[] = []
  for (const el of elements ?? []) {
    if (el.type === 'image' && el.rasterize) result.push(el)
    if ('children' in el && Array.isArray((el as any).children)) {
      result.push(...collectRasterizeImages((el as any).children))
    }
  }
  return result
}

function buildRasterizeImageJobs(slides: SlideData[]): SlideRasterizeJob[] {
  return slides.flatMap((s, i): SlideRasterizeJob[] => {
    const imgs = collectRasterizeImages(s.elements)
    if (imgs.length === 0) return []
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img): RasterizeTarget => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height),
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl
              delete img.rasterize
            },
          }),
        ),
      },
    ]
  })
}

/**
 * Rasterize partial-width background images (e.g. `![bg right:30%]`).
 *
 * When Marp uses split backgrounds, the <figure> element uses CSS
 * `background-size: cover` which may crop the image differently than PPTX's
 * default stretch-to-fill.  Screenshotting the rendered figure region gives
 * pixel-accurate reproduction.
 *
 * Only targets backgrounds that are NOT full-slide (partial width/height or
 * offset from origin) and have NOT already been rasterized by other jobs.
 */
function buildPartialBgJobs(slides: SlideData[]): SlideRasterizeJob[] {
  return slides.flatMap((s, i): SlideRasterizeJob[] => {
    const bgs = (s.backgroundImages ?? []).filter((b) => {
      // Skip already-rasterized backgrounds (cssFilter/cssFallback handled above)
      if (b.cssFilter || b.fromCssFallback) return false
      // Skip data: URLs — already embedded or rasterized
      if (b.url.startsWith('data:')) return false
      // Only rasterize partial-width/height backgrounds (split layouts)
      const isFullSlide =
        b.x <= 1 &&
        b.y <= 1 &&
        Math.abs(b.width - s.width) <= 2 &&
        Math.abs(b.height - s.height) <= 2
      return !isFullSlide
    })
    if (bgs.length === 0) return []
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg): RasterizeTarget => ({
            clip: {
              x: Math.round(bg.x),
              y: Math.round(bg.y),
              width: Math.round(bg.width),
              height: Math.round(bg.height),
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl
            },
          }),
        ),
        setup: (p) => hideAdvancedLayers(p),
        teardown: (p) => restoreAdvancedLayers(p),
      },
    ]
  })
}
