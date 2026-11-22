import type {
  BgImageData,
  SlideData,
  SlideElement,
  ImageElement,
  TextRun,
  TextStyle,
  ListItem,
  TableRow,
  TableCell,
  ElementBase,
  BadgeTextLine,
} from './types'
import {
  applyAutoScale,
  computeAutoScaleFactor,
  computeElementTransformScale,
  createDomSourceTracker,
} from './dom-reader'

/**
 * Extract structured slide data from the rendered DOM.
 *
 * **All helper functions are nested inside this function** so that Puppeteer's
 * `page.evaluate(extractSlides)` can serialise the entire function body in one
 * shot.  Because `page.evaluate` calls `Function.prototype.toString()` on the
 * argument, any references to module-scope variables would be lost after
 * serialisation.  Keeping everything inside a single closure avoids this
 * problem and remains safe even after webpack/esbuild minification (all
 * identifiers within the function body are renamed consistently).
 *
 * The function relies only on browser globals (`document`,
 * `getComputedStyle`, `Node`) and has no Node.js runtime imports.
 */
export function extractSlides(root: ParentNode = document): SlideData[] {
  const { sourceFor: sourceRefFor } = createDomSourceTracker()
  // -----------------------------------------------------------------
  // Helper: find effective background color for a slide section
  // -----------------------------------------------------------------
  function findBackgroundColor(section: Element): string {
    // 1. Check section's own background-color
    const style = getComputedStyle(section)
    const bg = style.backgroundColor
    if (bg && bg !== 'transparent' && bg !== 'rgba(0, 0, 0, 0)') {
      return bg
    }

    // 2. If background-color is transparent, try to extract a dominant color
    //    from CSS gradients in background-image
    const bgImage = style.backgroundImage
    if (bgImage && bgImage !== 'none') {
      const colorMatches = bgImage.match(
        /rgba?\(\s*\d+\s*,\s*\d+\s*,\s*\d+(?:\s*,\s*[\d.]+)?\s*\)/g,
      )
      if (colorMatches && colorMatches.length > 0) {
        // Use the last non-transparent color from the gradient stops
        for (let i = colorMatches.length - 1; i >= 0; i--) {
          const c = colorMatches[i]
          if (c !== 'rgba(0, 0, 0, 0)') {
            // Check alpha for rgba colors
            const alphaMatch = c.match(
              /rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*([\d.]+)\s*\)/,
            )
            if (!alphaMatch || parseFloat(alphaMatch[1]) > 0.1) {
              return c
            }
          }
        }
      }
    }

    // 3. Default to white — do NOT walk up to body/html because Marp's
    //    HTML uses black body background for the "between slides" area
    return 'rgb(255, 255, 255)'
  }

  // -----------------------------------------------------------------
  // Helper: extract text style from CSSStyleDeclaration
  // -----------------------------------------------------------------
  function extractTextStyle(
    style: CSSStyleDeclaration,
    includePadding = true,
  ): TextStyle {
    // Normalize CSS text-align: "start" → "left", "end" → "right"
    let textAlign = style.textAlign || 'left'
    if (textAlign === 'start') textAlign = 'left'
    else if (textAlign === 'end') textAlign = 'right'
    // Flex containers use justify-content for horizontal item alignment.
    // Map justify-content:center → textAlign:center so badge/label text is
    // horizontally centred in the PPTX text box.
    if (textAlign === 'left' && style.justifyContent === 'center')
      textAlign = 'center'

    return {
      color: style.color,
      fontSize: parseFloat(style.fontSize) || 16,
      fontFamily: style.fontFamily,
      fontWeight: parseInt(style.fontWeight, 10) || 400,
      textAlign: textAlign as TextStyle['textAlign'],
      lineHeight: parseFloat(style.lineHeight) || 0,
      letterSpacing: parseFloat(style.letterSpacing) || 0,
      ...(includePadding && parseFloat(style.paddingTop) > 0
        ? { paddingTop: parseFloat(style.paddingTop) }
        : {}),
      ...(includePadding && parseFloat(style.paddingRight) > 0
        ? { paddingRight: parseFloat(style.paddingRight) }
        : {}),
      ...(includePadding && parseFloat(style.paddingBottom) > 0
        ? { paddingBottom: parseFloat(style.paddingBottom) }
        : {}),
      ...(includePadding && parseFloat(style.paddingLeft) > 0
        ? { paddingLeft: parseFloat(style.paddingLeft) }
        : {}),
      ...(style.whiteSpace === 'nowrap' ? { whiteSpace: 'nowrap' as const } : {}),
    }
  }

  function extractStackingMetadata(style: CSSStyleDeclaration): {
    stackingIndex?: number
    stackingPositioned?: boolean
  } {
    const positioned =
      style.position === 'absolute' ||
      style.position === 'fixed' ||
      style.position === 'relative' ||
      style.position === 'sticky'
    const parsed = Number.parseInt(style.zIndex, 10)
    return {
      ...(positioned ? { stackingPositioned: true } : {}),
      ...(positioned && Number.isFinite(parsed)
        ? { stackingIndex: parsed }
        : {}),
    }
  }

  function isSemanticInlineHighlightTag(tag: string): boolean {
    return tag === 'strong' || tag === 'mark' || tag === 'code'
  }

  // -----------------------------------------------------------------
  // Helper: extract the effective background color from a computed style.
  //
  // Returns the solid backgroundColor when present; otherwise checks
  // backgroundImage for a two-stop linear-gradient where one stop is
  // transparent and the other is a solid color — the pattern used for
  // marker-style highlights (e.g. linear-gradient(transparent 62%, #fff2a8
  // 62%)).  Returns undefined when neither is set.
  //
  // This helper is used in both extractTextRuns (inline children of a
  // paragraph) and extractListItemEl (direct inline children of a tight
  // list <li>) so that the linear-gradient highlight path is not missed
  // when <strong> or similar elements appear as direct <li> children.
  // -----------------------------------------------------------------
  function extractEffectiveBg(elStyle: CSSStyleDeclaration): string | undefined {
    const bg = elStyle.backgroundColor
    const hasBg = !!bg && bg !== 'transparent' && bg !== 'rgba(0, 0, 0, 0)'
    if (hasBg) {
      const alphaMatch = bg.match(/,\s*([\d.]+)\s*\)$/)
      if (!alphaMatch || parseFloat(alphaMatch[1]) !== 0) return bg
    }
    const bi = elStyle.backgroundImage
    if (!bi || bi === 'none' || !bi.includes('linear-gradient')) return undefined
    const colorMatches = bi.match(/rgba?\([^)]+\)/g)
    if (!colorMatches || colorMatches.length === 0) return undefined
    for (let ci = colorMatches.length - 1; ci >= 0; ci--) {
      const c = colorMatches[ci]
      // Fast-path string checks before the regex: Chromium normalises
      // transparent to 'rgba(0, 0, 0, 0)' (with spaces) in most cases;
      // the regex is a safety net for other zero-alpha forms such as
      // 'rgba(0,0,0,0)' (no spaces) or 'rgba(r,g,b,0)' with non-zero rgb.
      // Alpha values are always integers in Chromium's computed output
      // (0 → '0', not '0.0'), so the regex uses \d+ for the alpha channel.
      if (
        c !== 'rgba(0, 0, 0, 0)' &&
        c !== 'rgba(0,0,0,0)' &&
        !/rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*0\s*\)/.test(c)
      ) {
        return c
      }
    }
    return undefined
  }

  // -----------------------------------------------------------------
  // Helper: detect emoji <img> elements
  //
  // Emoji libraries (e.g. Twemoji) replace emoji characters with <img>
  // elements. The original emoji character(s) are preserved in `alt`.
  // Detect by: explicit emoji class, Twemoji/emoji URL patterns, or
  // alt text that consists of Unicode Extended Pictographic characters.
  // -----------------------------------------------------------------
  function isEmojiImg(imgEl: HTMLImageElement): boolean {
    const alt = imgEl.alt ?? ''
    return !!(
      imgEl.classList?.contains('emoji') ||
      (imgEl.src &&
        (imgEl.src.includes('twemoji') || imgEl.src.includes('/emoji/'))) ||
      (alt.length > 0 &&
        alt.length <= 8 &&
        /\p{Extended_Pictographic}/u.test(alt))
    )
  }

  // -----------------------------------------------------------------
  // Helper: extract text runs from an element's child nodes.
  //
  // Design principles:
  //   - Inline elements (display:inline*) are flattened into the run    stream
  //     directly; their background-color is propagated to child runs.
  //   - Block-level elements (display:block/flex/grid/list-item/table) act as
  //     paragraph boundaries: a breakLine run is inserted between adjacent
  //     blocks so presentation software renders them on separate lines.
  //   - Text nodes containing '\n' are split at each newline so soft line-
  //     breaks in the source HTML (common in Marp blockquotes) become proper
  //     break runs instead of literal newline characters in outgoing text.
  //   - Trailing break runs are trimmed so callers receive a clean run list.
  // -----------------------------------------------------------------
  function extractTextRuns(
    element: Element,
    skipInlineBadges: boolean | Set<Element> = false,
    stripBgBadges: Set<Element> | false = false,
    skipBlockEls: Set<Element> | false = false,
  ): TextRun[] {
    const runs: TextRun[] = []

    const elementStyle = getComputedStyle(element)
    const elementBg = elementStyle.backgroundColor
    const elementHasBg =
      !(stripBgBadges instanceof Set && stripBgBadges.has(element)) &&
      !!elementBg &&
      elementBg !== 'transparent' &&
      elementBg !== 'rgba(0, 0, 0, 0)'

    function lastIsBreak(): boolean {
      return runs.length > 0 && runs[runs.length - 1].breakLine === true
    }

    // Push one or more text runs from a raw text string, splitting on '\n'
    // to convert soft line-breaks into explicit breakLine runs.
    function pushText(
      text: string,
      style: CSSStyleDeclaration,
      bg: string | undefined,
    ): void {
      const segments = text.split('\n')
      for (let i = 0; i < segments.length; i++) {
        const seg = segments[i]
        if (seg !== '') {
          const run: TextRun = {
            text: seg,
            color: style.color,
            fontSize: parseFloat(style.fontSize) || 16,
            fontFamily: style.fontFamily,
            bold: parseInt(style.fontWeight, 10) >= 600,
            italic: style.fontStyle === 'italic',
            underline: style.textDecorationLine?.includes('underline'),
            strikethrough: style.textDecorationLine?.includes('line-through'),
          }
          if (bg) run.backgroundColor = bg
          runs.push(run)
        }
        // Insert a break between segments, but never double-up
        if (i < segments.length - 1 && !lastIsBreak()) {
          runs.push({ text: '', breakLine: true })
        }
      }
    }

    for (const node of Array.from(element.childNodes)) {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? ''
        if (text.trim() === '') {
          // Preserve \n line breaks from whitespace-only text nodes.
          // These occur in <pre><code> between syntax-highlighted spans and
          // represent real line breaks (including blank lines).
          const newlineCount = (text.match(/\n/g) ?? []).length
          for (let i = 0; i < newlineCount; i++) {
            // Deduplicate: skip if the previous run is already a break.
            // Without this guard, a "<br>\n" sequence (br tag followed by
            // a whitespace text node) emits two consecutive breakLine runs,
            // which renders as a blank line in PPTX around emoji or block items.
            if (!lastIsBreak()) runs.push({ text: '', breakLine: true })
          }
          // Preserve inter-element spaces for whitespace-only nodes that contain
          // no newlines (e.g. the single space between syntax-highlighted keyword
          // spans such as "<span>async</span> <span>def</span>", or the space
          // between adjacent inline formatting elements like <em> and <strong>).
          // This applies to all text nodes with text.trim() === '' and zero
          // newlines, regardless of context (inline formatting or code blocks).
          // elementStyle is the computed style of the nearest parent Element.
          // For pre-formatted content (white-space: pre/pre-wrap/break-spaces),
          // preserve the exact whitespace so indentation is not collapsed. In
          // pre context this also means leading whitespace on the first run is
          // kept (runs may be empty), which is correct for code indentation.
          // pre-line collapses spaces/tabs per CSS spec (only preserves newlines),
          // so it is treated the same as normal and collapses to a single space.
          if (newlineCount === 0 && text.length > 0 && !lastIsBreak()) {
            const ws = elementStyle.whiteSpace
            const preserve = ws === 'pre' || ws === 'pre-wrap' || ws === 'break-spaces'
            pushText(preserve ? text : ' ', elementStyle, elementHasBg ? elementBg : undefined)
          }
          continue
        }
        pushText(text, elementStyle, elementHasBg ? elementBg : undefined)
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as Element
        const tag = el.tagName.toLowerCase()

        if (tag === 'br') {
          if (!lastIsBreak()) runs.push({ text: '', breakLine: true })
          continue
        }

        // KaTeX MathML accessibility duplicate: visually hidden via
        // position:absolute + clip:rect(1px,1px,1px,1px), NOT display:none.
        // Without this skip, extractTextRuns extracts text from both the hidden
        // .katex-mathml (MathML for screen readers) and the visible .katex-html,
        // doubling every math formula in the PPTX output.
        if (el.classList?.contains('katex-mathml')) continue

        // MathJax containers (<mjx-container class="MathJax">) wrap an <svg>
        // that renders the formula.  The SVG path data contains no extractable
        // text.  Skip here; extractNestedImages handles them as rasterized images.
        if (tag === 'mjx-container' || tag === 'svg') continue

        // Subscript / superscript elements: recurse and mark runs.
        // CSS position:relative + top/bottom offsets provide the visual shift
        // in the browser, but PPTX needs explicit subscript/superscript flags.
        if (tag === 'sub' || tag === 'sup') {
          const childRuns = extractTextRuns(el, skipInlineBadges, stripBgBadges)
          const flag = tag === 'sub' ? 'subscript' : 'superscript'
          childRuns.forEach((r) => {
            if (!r.breakLine) r[flag] = true
          })
          runs.push(...childRuns)
          continue
        }

        if (tag === 'a') {
          const href = (el as HTMLAnchorElement).href
          const childRuns = extractTextRuns(el)
          childRuns.forEach((r) => {
            if (!r.breakLine) r.hyperlink = href
          })
          runs.push(...childRuns)
          continue
        }

        if (tag === 'img') {
          const imgEl = el as HTMLImageElement
          const alt = imgEl.alt ?? ''
          if (isEmojiImg(imgEl) && alt) {
            pushText(alt, getComputedStyle(el), undefined)
          }
          continue
        }

        const elStyle = getComputedStyle(el)
        // Block-level elements act as paragraph separators
        if (/^(block|flex|grid|list-item|table)/.test(elStyle.display)) {
          // Caller can opt out specific block elements so they are extracted
          // separately (e.g. <ul>/<ol> inside <blockquote> emitted as list shapes).
          if (skipBlockEls instanceof Set && skipBlockEls.has(el)) continue
          if (!lastIsBreak() && runs.length > 0) {
            runs.push({ text: '', breakLine: true })
          }
          runs.push(...extractTextRuns(el))
        } else {
          // Inline element — recurse and propagate background-color.
          // Exception: inline-block/-flex/-grid elements with a non-transparent
          // background are rendered as separate badge/chip shapes by
          // extractInlineBadgeShapes, with their text rendered directly inside
          // the shape.  Skip them here to avoid duplicating the text in the
          // parent paragraph's text flow.
          const bg = elStyle.backgroundColor
          const hasBg = bg && bg !== 'transparent' && bg !== 'rgba(0, 0, 0, 0)'
          const alphaZero =
            hasBg &&
            (() => {
              const m = bg.match(/,\s*([\d.]+)\s*\)$/)
              return m ? parseFloat(m[1]) === 0 : false
            })()
          // When backgroundColor is transparent, check backgroundImage for a
          // two-stop linear-gradient where one stop is transparent and the
          // other is a solid color.  This pattern is used for marker-style
          // highlights (e.g. linear-gradient(transparent 62%, #fff2a8 62%)).
          // Extract the last non-transparent color stop as an approximate fill.
          const effectiveBg: string | undefined =
            stripBgBadges instanceof Set && stripBgBadges.has(el)
              ? undefined
              : extractEffectiveBg(elStyle)
          const inlineElBorderRadius = parseFloat(elStyle.borderRadius) || 0
          const inlineTag = el.tagName.toLowerCase()
          const isSemanticInlineHighlight = isSemanticInlineHighlightTag(inlineTag)
          // Structure-first badge detection: treat only inline <span> badges as
          // pill/chip candidates. Semantic inline text elements such as
          // <strong>, <mark>, and <code> keep their background on the text run
          // so they behave like marker-style highlights instead of detached
          // rounded shapes.
          const isInlineEligibleBadge = (() => {
            if (isSemanticInlineHighlight) return false
            if (elStyle.display !== 'inline') return false
            if (inlineTag !== 'span') return false
            if (inlineElBorderRadius <= 0) return false
            const m = (effectiveBg ?? '').match(/,\s*([\d.]+)\s*\)$/)
            return m ? parseFloat(m[1]) >= 0.5 : true // no alpha = fully opaque
          })()
          const isInlineBadgeDisplay =
            !isSemanticInlineHighlight &&
            (elStyle.display === 'inline-block' ||
              elStyle.display === 'inline-flex' ||
              elStyle.display === 'inline-grid')
          const isBadge =
            effectiveBg &&
            !alphaZero &&
            (isInlineBadgeDisplay || isInlineEligibleBadge)
          if (isBadge) {
            const shouldSkipBadge =
              skipInlineBadges === true ||
              (skipInlineBadges instanceof Set && skipInlineBadges.has(el))
            if (shouldSkipBadge) {
              // Isolated badge — text is rendered inside the badge shape, skip here
              continue
            }
            // Mixed badge (co-existing with other text).
            // When a background-only shape was emitted for this badge,
            // strip any backgroundColor from the runs (the round container shape
            // provides the visual); otherwise apply flat inline highlight.
            const shouldStripBg =
              stripBgBadges instanceof Set && stripBgBadges.has(el)
            const childRuns = extractTextRuns(el, false)
            if (shouldStripBg) {
              // Background-only shape handles visual; clear run bg to avoid flat overlay
              // (elementBg from recursive extractTextRuns already set bg on text runs)
              childRuns.forEach((r) => { if (!r.breakLine) r.backgroundColor = undefined })
            } else {
              childRuns.forEach((r) => {
                if (!r.breakLine && !r.backgroundColor) r.backgroundColor = effectiveBg
              })
            }
            runs.push(...childRuns)
            continue
          }
          const childRuns = extractTextRuns(el, skipInlineBadges, stripBgBadges)
          if (effectiveBg) {
            childRuns.forEach((r) => {
              if (!r.breakLine && !r.backgroundColor) r.backgroundColor = effectiveBg
            })
          }
          runs.push(...childRuns)
        }
      }
    }

    // Trim trailing break runs so callers get a clean list
    while (runs.length > 0 && runs[runs.length - 1].breakLine) {
      runs.pop()
    }
    // Trim leading break runs caused by HTML whitespace text nodes
    // (e.g. a newline after a block element's opening tag).  In HTML these do
    // not create visual space, but in PPTX a leading breakLine occupies a full
    // lineHeight, pushing the content down unexpectedly.
    while (runs.length > 0 && runs[0].breakLine) {
      runs.shift()
    }

    return runs
  }

  // -----------------------------------------------------------------
  // Helper: determine if an element has non-badge text content mixed with
  // inline badges.  Returns true when the element contains visible text
  // that is NOT exclusively from inline-block/-flex/-grid badge children.
  // -----------------------------------------------------------------
  // Helper: compute how far to shift a text element's left edge to clear
  // any inline badge shapes that sit flush at the container's left edge.
  // "Leading" badges are those whose x position is within 8 px of the
  // container's slide-relative left edge (tolerates border/padding offsets).
  //
  // For step-guide patterns like <h3><span.step>1</span>. heading</h3>,
  // the step badge sits at the heading's left → the heading text box is
  // shifted right by the badge width so it starts after the badge circle,
  // preventing textual overlap with the badge shape.
  // -----------------------------------------------------------------
  function computeLeadingOffset(
    badgeShapes: SlideElement[],
    containerRect: DOMRect,
    slideRect: DOMRect,
  ): number {
    if (badgeShapes.length === 0) return 0
    const containerSSLeft = containerRect.left - slideRect.left
    // Only full-container shapes (those with actual text runs) affect the leading
    // offset.  Background-only shapes (display:inline rounded bg, non-leading
    // inline-flex) have no runs and must not displace the paragraph text box.
    const leading = badgeShapes.filter((b) => {
      if (b.x > containerSSLeft + 8) return false
      const runs = (b as { runs?: unknown }).runs
      return Array.isArray(runs) && runs.length > 0
    })
    if (leading.length === 0) return 0
    const rightEdge = leading.reduce(
      (max, b) => Math.max(max, b.x + b.width),
      containerSSLeft,
    )
    return Math.max(0, rightEdge - containerSSLeft)
  }

  // -----------------------------------------------------------------
  // Helper: for shallow text recovery in flex/grid containers, shift the
  // recovered paragraph to start after the first visible child element.
  // This uses the actual DOM rect of the leading child so container padding
  // is handled correctly.
  // -----------------------------------------------------------------
  function computeShallowFlexOffset(
    container: Element,
    containerRect: DOMRect,
    slideRect: DOMRect,
    style: CSSStyleDeclaration,
  ): number {
    const containerSSLeft = containerRect.left - slideRect.left
    const containerGap = parseFloat(style.columnGap || style.gap || '0') || 0

    for (const node of Array.from(container.childNodes)) {
      if (node.nodeType === Node.TEXT_NODE) {
        if ((node.textContent ?? '').trim() !== '') return 0
        continue
      }
      if (node.nodeType !== Node.ELEMENT_NODE) continue

      const nodeEl = node as Element
      const nodeStyle = getComputedStyle(nodeEl)
      if (nodeStyle.display === 'none' || nodeStyle.visibility === 'hidden') {
        continue
      }

      const nodeRect = nodeEl.getBoundingClientRect()
      if (nodeRect.width <= 0 || nodeRect.height <= 0) continue

      return Math.max(0, nodeRect.right - slideRect.left - containerSSLeft) + containerGap
    }

    return 0
  }

  // -----------------------------------------------------------------
  // Helper: extract list items recursively
  // -----------------------------------------------------------------
  // Extract ListItem[] for a SINGLE <li> element at the given nesting level.
  // Exported as a helper so walkElements can call it per-<li> when splitting
  // a list around embedded images.
  // -----------------------------------------------------------------
  function extractListItemEl(
    li: Element,
    level = 0,
    skipBadges: Set<Element> | false = false,
    stripBadges: Set<Element> | false = false,
    leadingOffsetPx = 0,
  ): ListItem[] {
    const items: ListItem[] = []
    const runs: TextRun[] = []
    const nestedItems: ListItem[] = []

    function lastIsBreak(): boolean {
      return runs.length > 0 && runs[runs.length - 1].breakLine === true
    }

    for (const node of Array.from(li.childNodes)) {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? ''
        if (text.trim() === '') continue
        const liStyle = getComputedStyle(li)
        const liBg = liStyle.backgroundColor
        const liHasBg =
          !!liBg && liBg !== 'transparent' && liBg !== 'rgba(0, 0, 0, 0)'
        // Split on newlines so soft line-breaks in the markdown become breaks
        const segments = text.split('\n')
        for (let i = 0; i < segments.length; i++) {
          const seg = segments[i]
          if (seg !== '') {
            const run: TextRun = {
              text: seg,
              color: liStyle.color,
              fontSize: parseFloat(liStyle.fontSize) || 16,
              fontFamily: liStyle.fontFamily,
              bold: parseInt(liStyle.fontWeight, 10) >= 600,
              italic: liStyle.fontStyle === 'italic',
            }
            if (liHasBg) run.backgroundColor = liBg
            runs.push(run)
          }
          if (i < segments.length - 1) {
            if (!lastIsBreak()) runs.push({ text: '', breakLine: true })
          }
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as Element
        const childTag = el.tagName.toLowerCase()

        if (childTag === 'ul' || childTag === 'ol') {
          nestedItems.push(...extractListItems(el, level + 1))
        } else if (childTag === 'br') {
          // Markdown trailing-space hard line break inside a list item.
          // <br> is a direct child of <li> in tight lists; extractTextRuns
          // would not produce a break here because it is called on the <br>
          // element itself (no children).  Handle it explicitly.
          if (!lastIsBreak()) runs.push({ text: '', breakLine: true })
        } else if (childTag === 'img') {
          // Emoji img in a tight list item (no <p> wrapper).
          const imgEl = el as HTMLImageElement
          const alt = imgEl.alt ?? ''
          if (isEmojiImg(imgEl) && alt) {
            const liStyle = getComputedStyle(li)
            runs.push({
              text: alt,
              color: liStyle.color,
              fontSize: parseFloat(liStyle.fontSize) || 16,
              fontFamily: liStyle.fontFamily,
              bold: parseInt(liStyle.fontWeight, 10) >= 600,
              italic: liStyle.fontStyle === 'italic',
            })
          }
          // Non-emoji images are extracted by walkElements via
          // extractNestedImages; skip here to avoid duplicating them.
        } else if (childTag === 'sub' || childTag === 'sup') {
          // Subscript/superscript in tight list items: extract runs and mark them.
          const childRuns = extractTextRuns(el, skipBadges, stripBadges)
          const flag = childTag === 'sub' ? 'subscript' : 'superscript'
          childRuns.forEach((r) => { if (!r.breakLine) r[flag] = true })
          runs.push(...childRuns)
        } else {
          // Skip badge elements that were extracted as separate shapes by the
          // caller (walkElements passes a per-li Set via skipBadges).
          if (skipBadges !== false && skipBadges.has(el)) continue
          // For block-level children (e.g. <p> elements in a loose list where
          // markdown-it wraps each "paragraph" in <p>), insert a line break
          // between consecutive block elements.  Without this, loose list items
          // like <li><p>A</p><p>B</p></li> would have A and B merged into one
          // text run with no separator, losing the visual paragraph spacing.
          //
          // Use tag names (not getComputedStyle) so the check works correctly
          // in the jsdom test environment, where defaultStyles.display='block'
          // for all elements including inline ones like <strong>.
          const isBlockChild = /^(p|div|blockquote|pre|figure|h[1-6]|section|article|aside|header|footer|main)$/.test(childTag)
          if (isBlockChild && runs.length > 0 && !lastIsBreak()) {
            runs.push({ text: '', breakLine: true })
          }
          // <pre> / <marp-pre> inside <li> is extracted as a separate positioned
          // CodeElement by the <ul>/<ol> handler in walkElements — skip adding
          // code text as inline text runs here to avoid duplication and loss of
          // background fill.
          if (childTag === 'pre' || childTag === 'marp-pre') {
            continue
          }
          const childRuns = extractTextRuns(el, skipBadges, stripBadges)
          // When this element has a background-only shape (in stripBadges), strip
          // backgroundColor from its runs — the shaped provides the visual bg.
          if (stripBadges instanceof Set && stripBadges.has(el)) {
            childRuns.forEach((r) => { if (!r.breakLine) r.backgroundColor = undefined })
          } else if (!isBlockChild) {
            // For inline children (e.g. <strong> in a tight list), propagate the
            // effective background — including linear-gradient marker highlights —
            // to the child runs.  Block children (like <p> in a loose list) are
            // already handled inside extractTextRuns when it iterates their own
            // inline children, so they are excluded here.
            const elBg = extractEffectiveBg(getComputedStyle(el))
            if (elBg) {
              childRuns.forEach((r) => {
                if (!r.breakLine && !r.backgroundColor) r.backgroundColor = elBg
              })
            }
          }
          runs.push(...childRuns)
        }
      }
    }

    if (runs.length > 0) {
      const combinedText = runs.map((r) => r.text).join('')
      const liStyle = getComputedStyle(li)
      items.push({
        text: combinedText.trim(),
        level,
        listStyleType: liStyle.listStyleType || undefined,
        runs,
        ...(leadingOffsetPx > 0 ? { leadingOffset: leadingOffsetPx } : {}),
      })
    }
    items.push(...nestedItems)
    return items
  }

  // -----------------------------------------------------------------
  function extractListItems(
    list: Element,
    level = 0,
    perLiSkipMap?: Map<Element, Set<Element>>,
    perLiStripMap?: Map<Element, Set<Element>>,
    perLiLeadingOffsetMap?: Map<Element, number>,
  ): ListItem[] {
    const items: ListItem[] = []
    for (const child of Array.from(list.children)) {
      if (child.tagName.toLowerCase() === 'li') {
        const skipBadges = perLiSkipMap?.get(child) ?? false
        const stripBadges = perLiStripMap?.get(child) ?? false
        const leadingOffsetPx = perLiLeadingOffsetMap?.get(child) ?? 0
        items.push(
          ...extractListItemEl(
            child,
            level,
            skipBadges,
            stripBadges,
            leadingOffsetPx,
          ),
        )
      }
    }
    return items
  }

  // -----------------------------------------------------------------
  // Helper: extract syntax-highlighted code runs from a code element
  // -----------------------------------------------------------------
  function extractCodeRuns(codeEl: Element): TextRun[] {
    const runs: TextRun[] = []
    const defaultStyle = getComputedStyle(codeEl)

    function lastIsBreak(): boolean {
      return runs.length > 0 && runs[runs.length - 1].breakLine === true
    }

    function walk(node: Node) {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? ''
        if (text === '') return
        const parent = node.parentElement ?? codeEl
        const style = getComputedStyle(parent)
        // Split on newlines so that blank lines and line structure are preserved
        // as explicit breakLine runs rather than literal '\n' characters.
        const segments = text.split('\n')
        for (let i = 0; i < segments.length; i++) {
          const seg = segments[i]
          if (seg !== '') {
            runs.push({
              text: seg,
              color: style.color,
              fontSize:
                parseFloat(style.fontSize) ||
                parseFloat(defaultStyle.fontSize) ||
                16,
              fontFamily: style.fontFamily || defaultStyle.fontFamily,
              bold: parseInt(style.fontWeight, 10) >= 600,
              italic: style.fontStyle === 'italic',
            })
          }
          // Insert a break between segments (including blank lines).
          // Unlike extractTextRuns, do NOT suppress consecutive breaks here —
          // consecutive \n in code blocks represent blank lines that must be preserved.
          if (i < segments.length - 1) {
            runs.push({ text: '', breakLine: true })
          }
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        for (const child of Array.from(node.childNodes)) {
          walk(child)
        }
      }
    }

    walk(codeEl)
    return runs
  }

  // -----------------------------------------------------------------
  // Helper: extract table data with inline text runs
  // -----------------------------------------------------------------
  function extractTableData(table: Element): {
    rows: TableRow[]
    colWidths: number[]
  } {
    const rows: TableRow[] = []
    const colWidths: number[] = []

    for (const tr of Array.from(table.querySelectorAll('tr'))) {
      const cells: TableCell[] = []
      const isFirstRow = rows.length === 0
      // Extract row-level background colour.  CSS :nth-child alternating row
      // rules are applied to <tr>, so getComputedStyle(td).backgroundColor is
      // transparent even when the row appears visually coloured.  Fall back to
      // the <tr> colour when the individual cell is transparent.
      const trStyle = getComputedStyle(tr as Element)
      const trBg = trStyle.backgroundColor
      const trHasBg =
        !!trBg && trBg !== 'transparent' && trBg !== 'rgba(0, 0, 0, 0)'

      for (const td of Array.from(tr.querySelectorAll('th, td'))) {
        const style = getComputedStyle(td)
        const tdColSpan = (td as HTMLTableCellElement).colSpan || 1
        const tdRowSpan = (td as HTMLTableCellElement).rowSpan || 1
        if (isFirstRow) {
          // Capture rendered column widths from first row for proportional layout.
          // When a cell spans multiple columns, distribute its measured width evenly
          // so that colWidths has exactly one entry per grid column.
          const perColWidth = (td as HTMLElement).offsetWidth / tdColSpan
          for (let i = 0; i < tdColSpan; i++) {
            colWidths.push(perColWidth)
          }
        }
        const cellBg = style.backgroundColor
        const cellHasBg =
          !!cellBg && cellBg !== 'transparent' && cellBg !== 'rgba(0, 0, 0, 0)'
        const effectiveBg = cellHasBg ? cellBg : trHasBg ? trBg : cellBg
        const cellRuns = extractTextRuns(td)
        // Split runs at breakLine boundaries into structured paragraphs.
        // This mirrors OOXML's paragraph model and avoids relying on PptxGenJS
        // breakLine behaviour inside table cells.
        const paragraphs: { runs: TextRun[] }[] = [{ runs: [] }]
        for (const run of cellRuns) {
          if (run.breakLine) {
            paragraphs.push({ runs: [] })
          } else {
            paragraphs[paragraphs.length - 1].runs.push(run)
          }
        }
        // Remove trailing empty paragraph (no visible content)
        if (paragraphs.length > 1 && paragraphs[paragraphs.length - 1].runs.length === 0) {
          paragraphs.pop()
        }
        cells.push({
          text: td.textContent ?? '',
          runs: cellRuns,
          paragraphs,
          isHeader: td.tagName.toLowerCase() === 'th',
          ...(tdColSpan > 1 ? { colspan: tdColSpan } : {}),
          ...(tdRowSpan > 1 ? { rowspan: tdRowSpan } : {}),
          style: {
            color: style.color,
            backgroundColor: effectiveBg,
            fontSize: parseFloat(style.fontSize) || 16,
            fontFamily: style.fontFamily,
            fontWeight: parseInt(style.fontWeight, 10) || 400,
            textAlign: style.textAlign || 'left',
            borderColor: style.borderColor,
            borderWidth: parseFloat(style.borderTopWidth) || 0,
            ...(style.borderTopStyle && style.borderTopStyle !== 'none'
              ? { borderStyle: style.borderTopStyle }
              : {}),
            paddingTop: parseFloat(style.paddingTop) || 0,
            paddingRight: parseFloat(style.paddingRight) || 0,
            paddingBottom: parseFloat(style.paddingBottom) || 0,
            paddingLeft: parseFloat(style.paddingLeft) || 0,
          },
        })
      }

      rows.push({
        cells,
        ...(tr.getBoundingClientRect().height > 0
          ? { height: tr.getBoundingClientRect().height }
          : {}),
      })
    }

    return { rows, colWidths }
  }

  // -----------------------------------------------------------------
  // Helper: extract inline-block badge / chip shapes from a text container.
  //
  // Elements styled as `display: inline-block` with a non-transparent
  // background (e.g. `.step { display:inline-block; border-radius:999px;
  // background: var(--brand) }`) act as visual pill badges.  PPTX cannot
  // place a rounded-rectangle shape inside a text flow, so we extract the
  // badge as a separate positioned ContainerElement and emit it BEFORE the
  // parent text element so the text box sits on top.
  //
  // extractTextRuns intentionally omits propagating backgroundColor for
  // inline-block children so the text runs stay clean (no PPTX highlight);
  // the shape provides the background colour visually.
  // -----------------------------------------------------------------
  function extractBadgeTextLines(
    element: Element,
    slideRect: DOMRect,
  ): NonNullable<import('./types').ContainerElement['textLines']> | undefined {
    const lines: NonNullable<
      import('./types').ContainerElement['textLines']
    > = []
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
    let node: Node | null
    let measuredAnyCharacter = false

    while ((node = walker.nextNode())) {
      const parent = node.parentElement
      const text = node.textContent ?? ''
      if (!parent || !text) continue
      const style = getComputedStyle(parent)
      if (style.display === 'none' || style.visibility === 'hidden') continue
      let offset = 0
      for (const character of Array.from(text)) {
        const characterLength = character.length
        const range = document.createRange()
        if (
          typeof range.setStart !== 'function' ||
          typeof range.setEnd !== 'function' ||
          typeof range.getBoundingClientRect !== 'function'
        ) {
          return undefined
        }
        range.setStart(node, offset)
        range.setEnd(node, offset + characterLength)
        offset += characterLength
        const rect = range.getBoundingClientRect()
        if (rect.width <= 0 || rect.height <= 0) continue
        measuredAnyCharacter = true
        const topTolerance = Math.max(
          0.5,
          (parseFloat(style.lineHeight) || rect.height) * 0.2,
        )
        let line = lines.find(
          (candidate) => Math.abs(candidate.y - (rect.top - slideRect.top)) <= topTolerance,
        )
        if (!line) {
          line = {
            x: rect.left - slideRect.left,
            y: rect.top - slideRect.top,
            width: rect.width,
            height: rect.height,
            runs: [],
            lineHeight: parseFloat(style.lineHeight) || undefined,
            textAlign: (style.textAlign || 'left') as
              | 'left'
              | 'center'
              | 'right'
              | 'justify',
          }
          lines.push(line)
        } else {
          const right = line.x + line.width
          const bottom = line.y + line.height
          line.x = Math.min(line.x, rect.left - slideRect.left)
          line.y = Math.min(line.y, rect.top - slideRect.top)
          line.width =
            Math.max(right, rect.right - slideRect.left) - line.x
          line.height =
            Math.max(bottom, rect.bottom - slideRect.top) - line.y
        }
        line.runs.push({
          text: character,
          color: style.color,
          fontSize: parseFloat(style.fontSize) || 16,
          fontFamily: style.fontFamily,
          bold: parseInt(style.fontWeight, 10) >= 600,
          italic: style.fontStyle === 'italic',
          underline: style.textDecorationLine?.includes('underline'),
          strikethrough: style.textDecorationLine?.includes('line-through'),
        })
      }
    }

    for (const image of Array.from(
      element.querySelectorAll('img'),
    ) as HTMLImageElement[]) {
      if (!isEmojiImg(image)) continue
      const rect = image.getBoundingClientRect()
      if (rect.width <= 0 || rect.height <= 0) continue
      const imageTop = rect.top - slideRect.top
      const line = lines.find(
        (candidate) =>
          Math.abs(candidate.y - imageTop) <=
          Math.max(candidate.lineHeight ?? rect.height, rect.height) * 0.5,
      )
      if (!line || !image.alt) continue
      line.width =
        Math.max(line.x + line.width, rect.right - slideRect.left) - line.x
      line.height =
        Math.max(line.y + line.height, rect.bottom - slideRect.top) - line.y
      line.runs.push({
        text: image.alt,
        color: getComputedStyle(image).color,
        fontSize: parseFloat(getComputedStyle(image).fontSize) || 16,
        fontFamily: getComputedStyle(image).fontFamily,
      })
      measuredAnyCharacter = true
    }

    return measuredAnyCharacter
      ? lines.sort((left, right) => left.y - right.y)
      : undefined
  }

  function extractMeasuredParagraphLayout(
    element: Element,
    fallbackRuns: TextRun[],
    slideRect: DOMRect,
  ): {
    runs: TextRun[]
    lines: BadgeTextLine[]
  } | undefined {
    if (element.querySelector('img,svg,[data-bespoke-marp-fragment]'))
      return undefined
    const unsupportedBlock = Array.from(element.children).some((child) =>
      /^(block|flex|grid|list-item|table)/.test(
        getComputedStyle(child).display,
      ),
    )
    if (unsupportedBlock) return undefined

    const fallbackText = fallbackRuns
      .filter((run) => !run.breakLine)
      .map((run) => run.text)
      .join('')
    if (
      !/[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff01-\uff9f]/.test(
        fallbackText,
      )
    )
      return undefined
    const lines: Array<
      BadgeTextLine & {
        runs: Array<TextRun & { __styleKey?: string }>
      }
    > = []
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
    let node: Node | null

    while ((node = walker.nextNode())) {
      const text = node.textContent ?? ''
      const parent = node.parentElement
      if (!parent || !text) continue
      const style = getComputedStyle(parent)
      let offset = 0
      for (const character of Array.from(text)) {
        const charLength = character.length
        const range = document.createRange()
        if (
          typeof range.setStart !== 'function' ||
          typeof range.setEnd !== 'function' ||
          typeof range.getBoundingClientRect !== 'function'
        )
          return undefined
        range.setStart(node, offset)
        range.setEnd(node, offset + charLength)
        offset += charLength
        const rect = range.getBoundingClientRect()
        if (rect.width <= 0 || rect.height <= 0) continue
        let line = lines[lines.length - 1]
        if (
          !line ||
          Math.abs(rect.top - line.y) >
            Math.max(1, rect.height * 0.25)
        ) {
          line = {
            x: rect.left,
            y: rect.top,
            width: rect.width,
            height: rect.height,
            runs: [],
            lineHeight: parseFloat(style.lineHeight) || rect.height,
          }
          lines.push(line)
        } else {
          line.width =
            Math.max(line.x + line.width, rect.right) - line.x
          line.height = Math.max(line.height, rect.bottom - line.y)
        }
        const last = line.runs[line.runs.length - 1]
        const styleKey =
          `${style.color}|${style.fontSize}|${style.fontFamily}|` +
          `${style.fontWeight}|${style.fontStyle}|` +
          `${(parent.closest('a') as HTMLAnchorElement | null)?.href ?? ''}|` +
          `${extractEffectiveBg(style) ?? ''}`
        if (
          last && last.__styleKey === styleKey
        ) {
          last.text += character
        } else {
          line.runs.push({
            text: character,
            color: style.color,
            fontSize: parseFloat(style.fontSize) || 16,
            fontFamily: style.fontFamily,
            bold: parseInt(style.fontWeight, 10) >= 600,
            italic: style.fontStyle === 'italic',
            underline: style.textDecorationLine?.includes('underline'),
            strikethrough: style.textDecorationLine?.includes('line-through'),
            hyperlink:
              (parent.closest('a') as HTMLAnchorElement | null)?.href ||
              undefined,
            backgroundColor: extractEffectiveBg(style),
            subscript:
              parent.tagName.toLowerCase() === 'sub' ||
              parent.closest('sub') !== null ||
              undefined,
            superscript:
              parent.tagName.toLowerCase() === 'sup' ||
              parent.closest('sup') !== null ||
              undefined,
            __styleKey: styleKey,
          })
        }
      }
    }

    const measuredText = lines
      .flatMap((line) => line.runs)
      .map((run) => run.text)
      .join('')
    if (measuredText !== fallbackText) return undefined
    const cleanLines = lines.map((line) => {
      const runs = line.runs.map((run) => {
        delete (run as TextRun & { __styleKey?: string }).__styleKey
        return run
      })
      const align = fallbackRuns.length > 0
        ? getComputedStyle(element).textAlign
        : undefined
      return {
        x: line.x - slideRect.left,
        y: line.y - slideRect.top,
        width: line.width,
        height: line.height,
        runs,
        lineHeight: line.lineHeight,
        ...(align === 'center' || align === 'right' || align === 'justify'
          ? {
              textAlign: align as
                'center' |
                'right' |
                'justify',
            }
          : {}),
      }
    })
    const runs = cleanLines.flatMap((line, index) => [
      ...line.runs,
      ...(index < cleanLines.length - 1
        ? [{ text: '', breakLine: true as const }]
        : []),
    ])
    return { runs, lines: cleanLines }
  }

  function extractMeasuredParagraphRuns(
    element: Element,
    fallbackRuns: TextRun[],
    slideRect: DOMRect,
  ): TextRun[] | undefined {
    return extractMeasuredParagraphLayout(element, fallbackRuns, slideRect)?.runs
  }

  function extractInlineDecorationShapes(
    container: Element,
    slideRect: DOMRect,
  ): { shapes: SlideElement[]; elements: Element[] } {
    const shapes: SlideElement[] = []
    const elements: Element[] = []
    for (const element of Array.from(
      container.querySelectorAll('strong,mark,code,span'),
    )) {
      const style = getComputedStyle(element)
      const tag = element.tagName.toLowerCase()
      const hasGradient = style.backgroundImage.includes('linear-gradient')
      if (
        (tag === 'code' || tag === 'span') &&
        !hasGradient
      )
        continue
      const backgroundColor = extractEffectiveBg(style)
      if (!backgroundColor) continue
      if (
        tag === 'span' &&
        style.display !== 'inline'
      )
        continue
      const range = document.createRange()
      range.selectNodeContents(element)
      if (typeof range.getBoundingClientRect !== 'function') continue
      const rect = range.getBoundingClientRect()
      if (rect.width <= 0 || rect.height <= 0) continue
      const gradient = style.backgroundImage
      const markerStops = Array.from(
        gradient.matchAll(/(\d+(?:\.\d+)?)%/g),
      ).map((match) => Number(match[1]))
      const markerStop =
        gradient.includes('linear-gradient') && markerStops.length > 0
          ? Math.min(90, ...markerStops)
          : 0
      const y = rect.top - slideRect.top + rect.height * markerStop / 100
      const borderRadius = parseFloat(style.borderRadius) || 0
      shapes.push({
        type: 'container',
        children: [],
        x: rect.left - slideRect.left,
        y,
        width: rect.width,
        height: Math.max(1, rect.height * (1 - markerStop / 100)),
        style: {
          backgroundColor,
          ...(borderRadius > 0 ? { borderRadius } : {}),
        },
        source: sourceRefFor(element),
      })
      elements.push(element)
    }
    return { shapes, elements }
  }

  function extractInlineBadgeShapes(
    container: Element,
    slideRect: DOMRect,
    containerRect?: DOMRect,
  ): { shapes: SlideElement[]; elements: Element[]; bgOnlyElements: Element[] } {
    const badges: SlideElement[] = []
    const badgeEls: Element[] = []
    const bgOnlyElements: Element[] = []
    const containerSSLeft = containerRect
      ? containerRect.left - slideRect.left
      : -Infinity
    // Determine if the container has visible non-badge text content.
    // "Non-badge" here means: a non-empty TEXT_NODE or any child element that
    // is NOT display:inline-block/flex/grid (e.g. <strong>, plain <span>, text).
    //
    // This drives whether the inline-flex leading-only filter is applied:
    //   badge-only  (<p>HIGH MED LOW</p>) →  all badges extracted as shapes
    //   mixed       (<p>Install ② Create<p>) → leading-only filter to avoid
    //                                          mid-line shapes breaking flow
    const containerHasNonBadgeText = (() => {
      for (const node of Array.from(container.childNodes)) {
        if (node.nodeType === Node.TEXT_NODE) {
          if ((node.textContent ?? '').trim() !== '') return true
          continue
        }

        if (node.nodeType !== Node.ELEMENT_NODE) continue
        const cs = getComputedStyle(node as Element)
        if (cs.display === 'none' || cs.visibility === 'hidden') continue
        const isChildBadge =
          cs.display === 'inline-block' ||
          cs.display === 'inline-flex' ||
          cs.display === 'inline-grid'
        if (!isChildBadge) return true // non-badge element (strong, em, code…)
      }
      return false
    })()
    for (const el of Array.from(container.querySelectorAll('*'))) {
      const s = getComputedStyle(el as Element)
      // Match inline-block/flex/grid badges (classic badge pattern) OR
      // plain inline elements that have both borderRadius and a visible
      // background — these are visually rounded badges in HTML but PPTX
      // text highlight cannot render rounded corners.
      const inlineTag = (el as HTMLElement).tagName.toLowerCase()
      const isSemanticInlineHighlight = isSemanticInlineHighlightTag(inlineTag)
      const isInlineBadgeDisplay =
        !isSemanticInlineHighlight &&
        (s.display === 'inline-block' ||
          s.display === 'inline-flex' ||
          s.display === 'inline-grid')
      const inlineBorderRadius = parseFloat(s.borderRadius) || 0
      // Structure-first badge detection: only inline <span> elements are
      // treated as rounded badge candidates. Semantic inline text elements such
      // as <strong>, <mark>, and <code> remain text highlights even if they use
      // background-color + border-radius in CSS.
      const isInlineWithRoundedBg =
        !isSemanticInlineHighlight &&
        s.display === 'inline' &&
        inlineTag === 'span' &&
        inlineBorderRadius > 0
      if (!isInlineBadgeDisplay && !isInlineWithRoundedBg) continue
      const bg = s.backgroundColor
      if (!bg || bg === 'transparent') continue
      // Reject rgba() with alpha === 0 — handles both 'rgba(0, 0, 0, 0)' and
      // 'rgba(0,0,0,0)' (browser formatting varies).
      const alphaMatch = bg.match(/,\s*([\d.]+)\s*\)$/)
      if (alphaMatch && parseFloat(alphaMatch[1]) === 0) continue
      // For display:inline candidates, also reject semi-transparent backgrounds.
      // Inline <code> uses rgba with alpha ~0.06–0.12; real badge/highlight spans
      // use fully opaque colors (no alpha component).
      if (isInlineWithRoundedBg && alphaMatch && parseFloat(alphaMatch[1]) < 0.5) continue
      const iRect = (el as HTMLElement).getBoundingClientRect()
      if (iRect.width === 0 || iRect.height === 0) continue
      // For inline-block/flex/grid badges, apply a leading-only filter ONLY
      // when the container has surrounding non-badge text (mixed content).
      //
      // Mixed content (e.g. "Install ② Create config ✅ Verify"):
      //   Only LEADING badges (within 8 px of container left) become shapes.
      //   Non-leading badges stay as inline highlights so text flow is intact.
      //
      // Badge-only content (e.g. <p>HIGH MED LOW</p>):
      //   All badges are extracted as shapes regardless of position so that
      //   rounded corners are preserved in PPTX.
      if (
        isInlineBadgeDisplay &&
        containerRect &&
        containerHasNonBadgeText &&
        iRect.left - slideRect.left > containerSSLeft + 8
      ) {
        // Text around every badge is emitted as measured inline segments.
        // Treating all badges as shape+text preserves exact gaps instead of
        // asking PowerPoint to reflow text around empty overlay shapes.
      }
      // extractTextRuns receives the badge elements in skipInlineBadges and
      // omits their text from the parent flow, preventing duplication.
      const br = inlineBorderRadius
      // Capture badge text so it can be rendered directly inside the shape.
      const badgeRuns = extractTextRuns(el as Element)
      const textLines = extractBadgeTextLines(el as Element, slideRect)
      let textBounds:
        | (ElementBase & { lineHeight?: number })
        | undefined
      const textRange = document.createRange()
      textRange.selectNodeContents(el)
      if (typeof textRange.getBoundingClientRect === 'function') {
        const textRect = textRange.getBoundingClientRect()
        if (
          textRect.left < iRect.left - 0.5 ||
          textRect.right > iRect.right + 0.5 ||
          textRect.top < iRect.top - 0.5 ||
          textRect.bottom > iRect.bottom + 0.5
        ) {
          textBounds = {
            x: textRect.left - slideRect.left,
            y: textRect.top - slideRect.top,
            width: Math.max(1, textRect.width),
            height: Math.max(1, textRect.height),
            lineHeight: parseFloat(s.lineHeight) || undefined,
          }
        }
      }
      // Strip backgroundColor from badge runs: the container shape provides the
      // visual background.  Keeping highlight on the text creates a visible
      // artefact (the highlight bleeds outside the shape) when font metrics
      // cause a slight positioning mismatch between the shape and the text box.
      badgeRuns.forEach((r) => {
        if (!r.breakLine) r.backgroundColor = undefined
      })
      const hasBadgeText = badgeRuns.some(
        (r) => !r.breakLine && r.text.trim() !== '',
      )
      badges.push({
        type: 'container',
        children: [],
        ...(hasBadgeText ? { runs: badgeRuns } : {}),
        ...(textLines?.length ? { textLines } : {}),
        ...(textBounds ? { textBounds } : {}),
        x: iRect.left - slideRect.left,
        y: iRect.top - slideRect.top,
        width: iRect.width,
        height: iRect.height,
        style: {
          backgroundColor: bg,
          ...(br > 0 ? { borderRadius: br } : {}),
        },
      })
      badgeEls.push(el as Element)
    }
    return { shapes: badges, elements: badgeEls, bgOnlyElements }
  }

  function extractInlineTextSegments(
    container: Element,
    slideRect: DOMRect,
    badgeElements: Set<Element>,
    badgeShapesByElement?: Map<Element, SlideElement>,
  ): SlideElement[] {
    const segments: SlideElement[] = []

    function walk(node: Node): void {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as Element
        if (
          element.tagName.toLowerCase() === 'svg' ||
          element.tagName.toLowerCase() === 'mjx-container'
        ) {
          return
        }
        const badgeShape = badgeShapesByElement?.get(element)
        if (badgeShape) {
          segments.push(badgeShape)
          return
        }
        if (badgeElements.has(element)) return
        for (const child of Array.from(element.childNodes)) walk(child)
        return
      }
      if (node.nodeType !== Node.TEXT_NODE) return
      const parent = node.parentElement
      const text = node.textContent ?? ''
      if (!parent || text.trim() === '') return

      const range = document.createRange()
      range.selectNodeContents(node)
      if (typeof range.getBoundingClientRect !== 'function') return
      const rect = range.getBoundingClientRect()
      if (rect.width <= 0 || rect.height <= 0) return
      const style = getComputedStyle(parent)
      const source = sourceRefFor(parent)
      segments.push({
        type: 'paragraph',
        source: {
          ...source,
          id: `${source.id}:text`,
        },
        x: rect.left - slideRect.left,
        y: rect.top - slideRect.top,
        width: rect.width,
        height: rect.height,
        runs: [
          {
            text: text.trim(),
            color: style.color,
            fontSize: parseFloat(style.fontSize) || 16,
            fontFamily: style.fontFamily,
            bold: parseInt(style.fontWeight, 10) >= 600,
            italic: style.fontStyle === 'italic',
            underline: style.textDecorationLine?.includes('underline'),
            strikethrough: style.textDecorationLine?.includes('line-through'),
          },
        ],
        style: extractTextStyle(style, false),
      })
    }

    for (const child of Array.from(container.childNodes)) walk(child)
    return segments
  }

  // -----------------------------------------------------------------
  // Helper: collect IMG descendants of a container as ImageElements.
  // Called after processing text-bearing block elements (paragraph,
  // heading, blockquote, list, table) whose content is handled by
  // extractTextRuns / extractListItems, which skip <img> tags.
  // -----------------------------------------------------------------
  function resolveClippedRect(
    element: Element,
    sourceRect: DOMRect,
    slideRect: DOMRect,
  ): {
    left: number
    top: number
    width: number
    height: number
    isClipped: boolean
  } {
    let left = sourceRect.left
    let top = sourceRect.top
    let right = sourceRect.right
    let bottom = sourceRect.bottom
    let ancestor = element.parentElement
    let isClipped = false

    while (ancestor && ancestor.getBoundingClientRect) {
      const ancestorStyle = getComputedStyle(ancestor)
      const clips =
        ancestorStyle.overflow === 'hidden' ||
        ancestorStyle.overflow === 'clip' ||
        ancestorStyle.overflowX === 'hidden' ||
        ancestorStyle.overflowX === 'clip' ||
        ancestorStyle.overflowY === 'hidden' ||
        ancestorStyle.overflowY === 'clip'
      if (clips) {
        const ancestorRect = ancestor.getBoundingClientRect()
        const nextLeft = Math.max(left, ancestorRect.left)
        const nextTop = Math.max(top, ancestorRect.top)
        const nextRight = Math.min(right, ancestorRect.right)
        const nextBottom = Math.min(bottom, ancestorRect.bottom)
        isClipped ||= nextLeft !== left ||
          nextTop !== top ||
          nextRight !== right ||
          nextBottom !== bottom
        left = nextLeft
        top = nextTop
        right = nextRight
        bottom = nextBottom
      }
      if (ancestor === document.documentElement) break
      ancestor = ancestor.parentElement
    }

    const slideLeft = Math.max(left, slideRect.left)
    const slideTop = Math.max(top, slideRect.top)
    const slideRight = Math.min(right, slideRect.right)
    const slideBottom = Math.min(bottom, slideRect.bottom)
    isClipped ||= slideLeft !== left ||
      slideTop !== top ||
      slideRight !== right ||
      slideBottom !== bottom

    return {
      left: slideLeft,
      top: slideTop,
      width: Math.max(0, slideRight - slideLeft),
      height: Math.max(0, slideBottom - slideTop),
      isClipped,
    }
  }

  function extractNestedImages(
    el: Element,
    slideRect: DOMRect,
  ): ImageElement[] {
    const images: ImageElement[] = []
    for (const img of Array.from(el.querySelectorAll('img'))) {
      const imgEl = img as HTMLImageElement
      const rect = imgEl.getBoundingClientRect()
      if (rect.width === 0 || rect.height === 0) continue
      const s = getComputedStyle(imgEl)
      if (s.display === 'none' || s.visibility === 'hidden') continue
      // Skip emoji images — they are converted to text runs by extractTextRuns.
      // Use the same isEmojiImg() logic (class / src pattern / alt pictographic)
      // so the two paths stay consistent: an img skipped here must also be
      // handled (converted to alt text) by extractTextRuns.
      if (isEmojiImg(imgEl)) continue
      const cssFilter = s.filter && s.filter !== 'none' ? s.filter : undefined
      const clipped = resolveClippedRect(imgEl, rect, slideRect)
      if (clipped.width <= 0 || clipped.height <= 0) continue
      images.push({
        type: 'image',
        src: imgEl.src,
        naturalWidth: imgEl.naturalWidth,
        naturalHeight: imgEl.naturalHeight,
        x: clipped.left - slideRect.left,
        y: clipped.top - slideRect.top,
        width: clipped.width,
        height: clipped.height,
        ...(clipped.isClipped
          ? { rasterize: true }
          : cssFilter
            ? { cssFilter, pageX: rect.left, pageY: rect.top }
            : {}),
      })
    }

    // Inline SVGs (e.g. MathJax formulas, inline icons).
    // These cannot be rendered natively by PowerPoint; rasterize them as
    // screenshot images.  walkElements handles top-level SVGs separately;
    // this captures SVGs nested inside paragraphs, list items, etc.
    for (const svg of Array.from(el.querySelectorAll('svg'))) {
      const rect = svg.getBoundingClientRect()
      if (rect.width === 0 || rect.height === 0) continue
      const clipped = resolveClippedRect(svg, rect, slideRect)
      if (clipped.width <= 0 || clipped.height <= 0) continue
      try {
        const svgStr = new XMLSerializer().serializeToString(svg)
        const b64 = btoa(unescape(encodeURIComponent(svgStr)))
        const dataUrl = `data:image/svg+xml;base64,${b64}`
        images.push({
          type: 'image',
          src: dataUrl,
          naturalWidth: clipped.width,
          naturalHeight: clipped.height,
          x: clipped.left - slideRect.left,
          y: clipped.top - slideRect.top,
          width: clipped.width,
          height: clipped.height,
          rasterize: true,
        })
      } catch {
        // Skip if serialization fails
      }
    }

    return images
  }

  // -----------------------------------------------------------------
  // Helper: walk child elements and classify them
  // -----------------------------------------------------------------
  function extractCssBorders(style: CSSStyleDeclaration): {
    top?: { width: number; color: string; style?: string }
    right?: { width: number; color: string; style?: string }
    bottom?: { width: number; color: string; style?: string }
    left?: { width: number; color: string; style?: string }
  } {
    const side = (
      width: string,
      borderStyle: string,
      color: string,
    ): { width: number; color: string; style?: string } | undefined => {
      const parsedWidth = parseFloat(width) || 0
      if (parsedWidth <= 0 || borderStyle === 'none') return undefined
      return {
        width: parsedWidth,
        color,
        ...(borderStyle && borderStyle !== 'solid'
          ? { style: borderStyle }
          : {}),
      }
    }
    const top = side(style.borderTopWidth, style.borderTopStyle, style.borderTopColor)
    const right = side(style.borderRightWidth, style.borderRightStyle, style.borderRightColor)
    const bottom = side(style.borderBottomWidth, style.borderBottomStyle, style.borderBottomColor)
    const left = side(style.borderLeftWidth, style.borderLeftStyle, style.borderLeftColor)
    return {
      ...(top ? { top } : {}),
      ...(right ? { right } : {}),
      ...(bottom ? { bottom } : {}),
      ...(left ? { left } : {}),
    }
  }

  function extractAvailableFlowWidth(
    element: Element,
    rect: DOMRect,
  ): number {
    const parent = element.parentElement
    if (!parent) return rect.width
    const parentRect = parent.getBoundingClientRect()
    if (parentRect.width <= 0 || parentRect.right < rect.left) return rect.width

    const rightEdge = Array.from(parent.children)
      .filter((sibling) => sibling !== element)
      .map((sibling) => sibling.getBoundingClientRect())
      .filter(
        (siblingRect) =>
          siblingRect.width > 0 &&
          siblingRect.height > 0 &&
          siblingRect.left >= rect.right - 1 &&
          siblingRect.top < rect.bottom &&
          siblingRect.bottom > rect.top,
      )
      .reduce(
        (right, siblingRect) => Math.min(right, siblingRect.left),
        parentRect.right,
      )

    return Math.max(1, rightEdge - rect.left)
  }

  function walkElement(
    parent: Element,
    child: Element,
    slideRect: DOMRect,
  ): SlideElement[] {
    const elements: SlideElement[] = []
    const style = getComputedStyle(child)
    if (style.display === 'none') return []

    if ((child as HTMLElement).dataset?.marpitPresenterNotes !== undefined)
      return []

    // Skip Marp advanced background container (handled at slide level)
    if (
      (child as HTMLElement).dataset?.marpitAdvancedBackgroundContainer !==
      undefined
    )
      return []

    // display:contents removes the wrapper's own box while retaining its
    // children in the rendered layout. Recurse through all child nodes so text
    // nodes participate alongside element nodes in the same DOM order.
    if (style.display === 'contents') {
      return walkDisplayContents(child, parent, slideRect)
    }

    if (style.visibility === 'hidden') return []

    const rect = child.getBoundingClientRect()
    if (rect.width === 0 || rect.height === 0) return []

      const tag = child.tagName.toLowerCase()

      // Skip display:inline elements — their text content is captured by
      // extractTextRuns on the containing element.
      // <img> and <svg> are intentionally excluded: visual elements must be
      // captured even when their computed display is 'inline'.
      // In flex/grid containers, all direct children become block-like items
      // regardless of their display property, so inline children must also be
      // walked (otherwise flex-row text spans would be silently dropped).
      const parentIsFlexOrGrid = /^(flex|inline-flex|grid|inline-grid)/.test(
        getComputedStyle(parent).display,
      )
      if (
        !parentIsFlexOrGrid &&
        tag !== 'img' &&
        tag !== 'svg' &&
        style.display === 'inline'
      )
        return []

      const base = {
        source: sourceRefFor(child),
        x: rect.left - slideRect.left,
        y: rect.top - slideRect.top,
        width: rect.width,
        height: rect.height,
        ...extractStackingMetadata(style),
      }

      if (/^h[1-6]$/.test(tag)) {
        // Capture CSS border decorations (border-bottom for h1 rules, border-left
        // for h2 side bars). These are set by the theme's stylesheet and are not
        // visible in the DOM tree, only in computed styles.
        const borderBottomWidth = parseFloat(style.borderBottomWidth) || 0
        const borderBottomStyle = style.borderBottomStyle
        const borderLeftWidth = parseFloat(style.borderLeftWidth) || 0
        // Extract inline badge shapes (pill/circle steps, status chips, etc.).
        // Shapes are ALWAYS emitted so badges render as rounded shapes in PPTX.
        // For badges at the container's left edge ("leading badges", e.g. step
        // numbers like <span.step>1</span>. heading), the heading text box is
        // shifted right by the badge width to prevent textual overlap with the
        // badge shape.
        const { shapes: headingBadgeShapes, elements: headingBadgeEls, bgOnlyElements: headingBgOnlyEls } =
          extractInlineBadgeShapes(child, slideRect, rect)
        const headingLeadingOffset = computeLeadingOffset(
          headingBadgeShapes,
          rect,
          slideRect,
        )
        if (headingBadgeShapes.length > 0) elements.push(...headingBadgeShapes)
        const headingBadgeSet: boolean | Set<Element> =
          headingBadgeEls.length > 0 ? new Set(headingBadgeEls) : false
        const headingBgOnlySet: Set<Element> | false =
          headingBgOnlyEls.length > 0 ? new Set(headingBgOnlyEls) : false
        // Extract padding for headings — same pattern as blockquote.
        // paddingLeft provides the gap between the border-left bar and text;
        // other paddings are passed through for text inset in PPTX.
        const headingPaddingTop = parseFloat(style.paddingTop) || 0
        const headingPaddingRight = parseFloat(style.paddingRight) || 0
        const headingPaddingBottom = parseFloat(style.paddingBottom) || 0
        const headingPaddingLeft = parseFloat(style.paddingLeft) || 0
        const headingRuns = extractTextRuns(child, headingBadgeSet, headingBgOnlySet)
        // For isolated badges (no surrounding text), omit the empty heading.
        if (
          headingBadgeShapes.length === 0 ||
          headingRuns.some((r) => !r.breakLine && r.text.trim() !== '')
        ) {
          elements.push({
            type: 'heading',
            level: parseInt(tag[1], 10),
            runs: headingRuns,
            ...base,
            x: base.x + headingLeadingOffset,
            width: Math.max(10, base.width - headingLeadingOffset),
            style: {
              ...extractTextStyle(style),
              ...(headingPaddingTop || headingPaddingRight || headingPaddingBottom || headingPaddingLeft
                ? { paddingTop: headingPaddingTop, paddingRight: headingPaddingRight, paddingBottom: headingPaddingBottom, paddingLeft: headingPaddingLeft }
                : {}),
            },
            ...(borderBottomWidth > 0
              ? {
                  borderBottom: {
                    width: borderBottomWidth,
                    color: style.borderBottomColor,
                    ...(borderBottomStyle && borderBottomStyle !== 'solid' && borderBottomStyle !== 'none'
                      ? { style: borderBottomStyle }
                      : {}),
                  },
                }
              : {}),
            ...(borderLeftWidth > 0
              ? {
                  borderLeft: {
                    width: borderLeftWidth,
                    color: style.borderLeftColor,
                  },
                }
              : {}),
          })
        }
        elements.push(...extractNestedImages(child, slideRect))
      } else if (tag === 'p') {
        // When a paragraph contains inline MathJax formulas (mjx-container)
        // mixed with surrounding text, the formula SVGs and text cannot be
        // independently positioned in PPTX (text box would overlap with images).
        // Rasterize the entire paragraph as a single screenshot image.
        if (child.querySelector('mjx-container')) {
          const pRect = child.getBoundingClientRect()
          if (pRect.width > 0 && pRect.height > 0) {
            elements.push({
              type: 'image',
              src: '',  // placeholder — rasterizeSlideTargets fills this via screenshot
              naturalWidth: pRect.width,
              naturalHeight: pRect.height,
              x: pRect.left - slideRect.left,
              y: pRect.top - slideRect.top,
              width: pRect.width,
              height: pRect.height,
              rasterize: true,
            } as any)
          }
          return elements
        }
        // Extract inline badge shapes. Shapes are always emitted so badges
        // render as rounded pill/circle elements in PPTX.  For leading badges
        // (at the paragraph's left edge), the paragraph text box is shifted
        // right to avoid overlap with the badge shape.
        const { shapes: paraBadgeShapes, elements: paraBadgeEls, bgOnlyElements: paraBgOnlyEls } =
          extractInlineBadgeShapes(child, slideRect, rect)
        const {
          shapes: paraDecorationShapes,
          elements: paraDecorationEls,
        } = extractInlineDecorationShapes(child, slideRect)
        if (paraDecorationShapes.length > 0)
          elements.push(...paraDecorationShapes)
        const paraLeadingOffset = computeLeadingOffset(
          paraBadgeShapes,
          rect,
          slideRect,
        )
        // Detect leading inline images that affect paragraph positioning.
        //
        // Case A — image before text with no <br>:  "![w:300](img) caption text"
        //   → in HTML the text flows to the right of the image on the same line;
        //     shift the paragraph x so it starts after the image's right edge.
        //
        // Case B — image followed by <br> then text:  "![w:300](img)\ncaption"
        //   → Marp renders this as <p><img><br>caption</p>; the text appears
        //     below the image; shift the paragraph y down by the image height.
        let inlineImgXOffset = 0
        let inlineImgYOffset = 0
        {
          let firstNonEmojiImg: HTMLImageElement | null = null
          let seenBrAfterImg = false
          for (const node of Array.from(child.childNodes)) {
            if (node.nodeType === Node.ELEMENT_NODE) {
              const en = node as Element
              const enTag = en.tagName.toLowerCase()
              if (enTag === 'img') {
                const ie = en as HTMLImageElement
                if (!isEmojiImg(ie)) {
                  if (!firstNonEmojiImg) firstNonEmojiImg = ie
                }
                continue
              }
              if (enTag === 'br' && firstNonEmojiImg && !seenBrAfterImg) {
                seenBrAfterImg = true
                continue
              }
              break // other element ends the leading run
            } else if (node.nodeType === Node.TEXT_NODE) {
              if ((node.textContent ?? '').trim() !== '') break // real text — stop
              // whitespace-only node between <img> and <br> → keep scanning
            }
          }
          if (firstNonEmojiImg) {
            const imgR = firstNonEmojiImg.getBoundingClientRect()
            if (seenBrAfterImg) {
              // Case B: text is below the image
              inlineImgYOffset = imgR.bottom - rect.top
            } else {
              // Case A: text is beside the image.
              // CSS default vertical-align:baseline places the text baseline at the
              // image's bottom edge.  Move x to the right of the image, and move y
              // down so the text box starts roughly one line-height above the image
              // bottom — matching the visual "bottom-right" position in the browser.
              inlineImgXOffset = imgR.right - rect.left
              const parsedLH = parseFloat(style.lineHeight)
              const lineHeight =
                !isNaN(parsedLH) && parsedLH > 0
                  ? parsedLH
                  : (parseFloat(style.fontSize) || 16) * 1.5
              inlineImgYOffset = Math.max(0, (imgR.bottom - rect.top) - lineHeight)
            }
          }
        }

        const paraBadgeSet: boolean | Set<Element> =
          paraBadgeEls.length > 0 ? new Set(paraBadgeEls) : false
        const paraBgOnlySet: Set<Element> | false =
          paraBgOnlyEls.length > 0 || paraDecorationEls.length > 0
            ? new Set([...paraBgOnlyEls, ...paraDecorationEls])
            : false
        if (paraBadgeEls.length > 0) {
          const textSegments = extractInlineTextSegments(
            child,
            slideRect,
            new Set(paraBadgeEls),
            new Map(
              paraBadgeEls.map((element, index) => [
                element,
                paraBadgeShapes[index],
              ]),
            ),
          )
          const hasMeasuredText = textSegments.some(
            (segment) => segment.type === 'paragraph',
          )
          if (hasMeasuredText) {
            elements.push(...textSegments)
            elements.push(...extractNestedImages(child, slideRect))
            return elements
          }
          elements.push(...paraBadgeShapes)
        }
        {
          const extractedRuns = extractTextRuns(
            child,
            paraBadgeSet,
            paraBgOnlySet,
          )
          const measuredLayout =
            paraBadgeEls.length === 0 &&
            paraDecorationEls.length === 0
              ? extractMeasuredParagraphLayout(
                  child,
                  extractedRuns,
                  slideRect,
                )
              : undefined
          const runs =
            measuredLayout?.runs ?? extractedRuns
          // Only emit a paragraph if it has visible text; images are extracted below.
          if (runs.some((r) => !r.breakLine && r.text.trim() !== '')) {
            elements.push({
              type: 'paragraph',
              runs,
              ...base,
              x: base.x + paraLeadingOffset + inlineImgXOffset,
              y: base.y + inlineImgYOffset,
              width: Math.max(10, base.width - paraLeadingOffset - inlineImgXOffset),
              height: Math.max(10, base.height - inlineImgYOffset),
              style: extractTextStyle(style),
              ...(measuredLayout && measuredLayout.lines.length > 1
                ? { measuredLines: measuredLayout.lines }
                : {}),
            })
          }
        }
        elements.push(...extractNestedImages(child, slideRect))
      } else if (tag === 'ul' || tag === 'ol') {
        // List badges (e.g. <span class="badge"> inside <li>) are handled by
        // extractTextRuns called from extractListItems.  Badge text is kept in
        // the list run flow as inline highlights (backgroundColor) rather than
        // extracted as separate shapes, because badges inside list items are
        // always mixed with surrounding text.
        //
        // When a <li> contains an embedded non-emoji image (e.g. markdown
        // image syntax without blank lines inserted between list items), the
        // <ul> bounding box spans both the text items AND the image area.
        // Rendering the entire list as one PptxGenJS text box causes the list
        // items to stack at the top, overlapping the image that is extracted
        // separately via extractNestedImages.  Split the list into sub-lists
        // around each image-containing <li> so that each sub-list occupies
        // only the vertical space of its own <li> elements, with images
        // interleaved at their actual rendered positions.
        const liChildren = Array.from(child.children).filter(
          (c) => c.tagName.toLowerCase() === 'li',
        )
        const hasEmbeddedImage = liChildren.some((li) =>
          Array.from(li.querySelectorAll('img')).some(
            (img) => !isEmojiImg(img as HTMLImageElement),
          ),
        )
        const hasEmbeddedCode = liChildren.some((li) =>
          Array.from(li.children).some((c) => {
            const t = c.tagName.toLowerCase()
            return t === 'pre' || t === 'marp-pre'
          }),
        )

        // Read <ol start="N"> attribute; defaults to 1 for <ul> and unset <ol>.
        const olStart = tag === 'ol' ? ((child as HTMLOListElement).start || 1) : 1
        const olHasExplicitStart = tag === 'ol' && child.hasAttribute('start')

        if (!hasEmbeddedImage && !hasEmbeddedCode) {
          // Fast path: no embedded images.
          // Extract inline badge shapes from each <li> so that rounded-corner
          // badges (e.g. <span style="border-radius:8px;background:#c05621">)
          // are rendered as positioned shapes rather than flat text highlights,
          // matching the paragraph badge extraction behaviour.
          const liBadgeSets = new Map<Element, Set<Element>>()
          const liBgOnlySets = new Map<Element, Set<Element>>()
          const liLeadingOffsetMap = new Map<Element, number>()
          const liBadgeEntries = new Map<
            Element,
            Array<{ element: Element; shape: SlideElement }>
          >()
          for (const li of liChildren) {
            const liRect = li.getBoundingClientRect()
            const { shapes: liBadgeShapes, elements: liBadgeEls, bgOnlyElements: liBgOnlyEls } =
              extractInlineBadgeShapes(li, slideRect, liRect)
            const {
              shapes: liDecorationShapes,
              elements: liDecorationEls,
            } = extractInlineDecorationShapes(li, slideRect)
            if (liDecorationShapes.length > 0)
              elements.push(...liDecorationShapes)
            const directBadgeIndexes = liBadgeEls
              .map((element, index) => ({ element, index }))
              .filter(({ element }) => element.closest('li') === li)
            const directBadgeElements = directBadgeIndexes.map(
              ({ element }) => element,
            )
            const directBadgeShapes = directBadgeIndexes.map(
              ({ index }) => liBadgeShapes[index],
            )
            const liLeadingOffset = computeLeadingOffset(
              directBadgeShapes,
              liRect,
              slideRect,
            )
            if (directBadgeShapes.length > 0) {
              liBadgeEntries.set(
                li,
                directBadgeElements.map((element, index) => ({
                  element,
                  shape: directBadgeShapes[index],
                })),
              )
            }
            if (directBadgeElements.length > 0) {
              liBadgeSets.set(li, new Set(directBadgeElements))
            }
            if (
              liBgOnlyEls.length > 0 ||
              liDecorationEls.length > 0
            ) {
              liBgOnlySets.set(
                li,
                new Set(
                  [
                    ...liBgOnlyEls.filter(
                      (element) => element.closest('li') === li,
                    ),
                    ...liDecorationEls.filter(
                      (element) => element.closest('li') === li,
                    ),
                  ],
                ),
              )
            }
            if (liLeadingOffset > 0) {
              liLeadingOffsetMap.set(li, liLeadingOffset)
            }
          }
          if (liBadgeEntries.size === 0) {
            elements.push({
              type: 'list',
              ordered: tag === 'ol',
              ...(olHasExplicitStart ? { startNumber: olStart } : {}),
              listStyleType: style.listStyleType || undefined,
              items: extractListItems(child),
              ...base,
              availableWidth: extractAvailableFlowWidth(child, rect),
              style: extractTextStyle(style),
            })
          } else {
            let orderedIndex = olStart
            let pendingItems: ListItem[] = []
            let pendingTop = 0
            let pendingBottom = 0
            const flushPending = () => {
              if (pendingItems.length === 0) return
              elements.push({
                type: 'list',
                ordered: tag === 'ol',
                ...((olHasExplicitStart || orderedIndex > 1)
                  ? { startNumber: orderedIndex }
                  : {}),
                listStyleType: style.listStyleType || undefined,
                items: pendingItems,
                x: base.x,
                y: pendingTop,
                width: base.width,
                height: Math.max(10, pendingBottom - pendingTop),
                availableWidth: extractAvailableFlowWidth(child, rect),
                style: extractTextStyle(style),
              })
              orderedIndex += pendingItems.length
              pendingItems = []
            }
            for (const li of liChildren) {
              const liRect = li.getBoundingClientRect()
              const badgeEntries = liBadgeEntries.get(li)
              if (!badgeEntries) {
                const liItems = extractListItemEl(li)
                if (pendingItems.length === 0) {
                  pendingTop = liRect.top - slideRect.top
                }
                pendingBottom = liRect.bottom - slideRect.top
                pendingItems.push(...liItems)
                continue
              }

              flushPending()
              const flow = extractInlineTextSegments(
                li,
                slideRect,
                new Set(badgeEntries.map(({ element }) => element)),
                new Map(
                  badgeEntries.map(({ element, shape }) => [element, shape]),
                ),
              )
              if (!flow.some((element) => element.type === 'paragraph')) {
                const fallbackItems = extractListItemEl(
                  li,
                  0,
                  liBadgeSets.get(li) ?? false,
                  liBgOnlySets.get(li) ?? false,
                  liLeadingOffsetMap.get(li),
                )
                const fallbackList = () => {
                  if (fallbackItems.length === 0) return
                  const badgeLeft = Math.min(
                    ...badgeEntries.map(({ shape }) => shape.x),
                  )
                  elements.push({
                    type: 'list',
                    ordered: tag === 'ol',
                    ...((olHasExplicitStart || orderedIndex > 1)
                      ? { startNumber: orderedIndex }
                      : {}),
                    listStyleType: style.listStyleType || undefined,
                    items: fallbackItems,
                    x: base.x,
                    y: liRect.top - slideRect.top,
                    width: Math.max(
                      10,
                      Math.min(
                        liRect.width,
                        badgeLeft -
                          (liRect.left - slideRect.left) -
                          2,
                      ),
                    ),
                    height: liRect.height,
                    availableWidth: extractAvailableFlowWidth(child, rect),
                    style: extractTextStyle(style),
                  })
                  orderedIndex += 1
                }
                if ((liLeadingOffsetMap.get(li) ?? 0) > 0) {
                  elements.push(...badgeEntries.map(({ shape }) => shape))
                  fallbackList()
                } else {
                  fallbackList()
                  elements.push(...badgeEntries.map(({ shape }) => shape))
                }
                continue
              }
              let listPrefixSegments: Array<
                Extract<SlideElement, { type: 'paragraph' }>
              > = []
              let listPrefixEmitted = false
              const emitPrefixList = () => {
                if (
                  listPrefixEmitted ||
                  listPrefixSegments.length === 0
                )
                  return
                const badgeLeft = Math.min(
                  ...badgeEntries.map(({ shape }) => shape.x),
                )
                const listStyle = extractTextStyle(style)
                const firstSegment = listPrefixSegments[0]
                elements.push({
                  type: 'list',
                  ordered: tag === 'ol',
                  ...(tag === 'ol' ? { startNumber: orderedIndex } : {}),
                  listStyleType: style.listStyleType || undefined,
                  items: [{
                    text: '\u00a0',
                    level: 0,
                    listStyleType: style.listStyleType || undefined,
                    runs: [{
                      text: '\u00a0',
                      color: listStyle.color,
                      fontSize: listStyle.fontSize,
                      fontFamily: listStyle.fontFamily,
                      bold: listStyle.fontWeight >= 600,
                    }],
                  }],
                  x: base.x,
                  y: firstSegment.y,
                  width: Math.max(
                    10,
                    firstSegment.x -
                      base.x -
                      2 +
                      40,
                  ),
                  height: Math.max(firstSegment.height, liRect.height),
                  style: listStyle,
                })
                elements.push(
                  ...listPrefixSegments.map((segment) => ({
                    ...segment,
                    width: Math.max(
                      10,
                      badgeLeft - segment.x - 2,
                    ),
                  })),
                )
                listPrefixEmitted = true
                orderedIndex++
              }
              for (const flowElement of flow) {
                if (flowElement.type === 'container') {
                  emitPrefixList()
                  elements.push(flowElement)
                } else if (
                  flowElement.type === 'paragraph' &&
                  !listPrefixEmitted
                ) {
                  listPrefixSegments.push(flowElement)
                } else {
                  elements.push(flowElement)
                }
              }
              emitPrefixList()
            }
            flushPending()
          }
          elements.push(...extractNestedImages(child, slideRect))
          // Extract direct <pre> children of each <li> as separate CodeElement
          // shapes, positioned at the <pre> element's actual rendered location.
          // The <pre> content is deliberately NOT included in the list item
          // text runs (see extractListItemEl) so the code block retains its
          // background fill rather than appearing as character-level highlights.
          for (const li of liChildren) {
            for (const preChild of Array.from(li.children)) {
              const preChildTag = preChild.tagName.toLowerCase()
              if (preChildTag !== 'pre' && preChildTag !== 'marp-pre') continue
              if (preChild.querySelector('svg')) continue // handled as image
              const preRect = preChild.getBoundingClientRect()
              if (preRect.width <= 0 || preRect.height <= 0) continue
              const preStyle = getComputedStyle(preChild)
              const codeEl = preChild.querySelector('code')
              const codeTarget = codeEl ?? preChild
              // Resolve background: prefer <pre>, fall back to <code>
              const preBg = preStyle.backgroundColor
              const codeBg = (preBg && preBg !== 'transparent' && preBg !== 'rgba(0, 0, 0, 0)' && preBg !== 'rgba(0,0,0,0)')
                ? preBg
                : (codeEl ? getComputedStyle(codeEl).backgroundColor : preBg)
              const preCodeRuns = extractCodeRuns(codeTarget)
              const preScaleFactor = computeAutoScaleFactor(
                preCodeRuns,
                preStyle,
                preRect.height,
                computeElementTransformScale(preChild),
              )
              const preBaseStyle = extractTextStyle(preStyle)
              const preAdjustedFs = applyAutoScale(preCodeRuns, preBaseStyle.fontSize, preScaleFactor)
              elements.push({
                type: 'code',
                text: codeTarget.textContent ?? '',
                language: codeEl?.className?.replace('language-', '') ?? '',
                runs: preCodeRuns,
                x: preRect.left - slideRect.left,
                y: preRect.top - slideRect.top,
                width: preRect.width,
                height: preRect.height,
                style: {
                  ...preBaseStyle,
                  fontSize: preAdjustedFs,
                  backgroundColor: codeBg,
                },
              })
            }
          }
        } else {
          // Split the list around image-containing <li> elements so the
          // extracted images are interleaved at their actual positions.
          // olRunningNumber tracks the start number for each sub-list chunk so
          // that ordered list numbering continues correctly across image breaks.
          let olRunningNumber = olStart
          let pendingItems: ListItem[] = []
          let pendingTop = -1
          let pendingBottom = -1

          const flushPending = () => {
            if (pendingItems.length === 0) return
            const chunkStart = olRunningNumber
            elements.push({
              type: 'list',
              ordered: tag === 'ol',
              ...(olHasExplicitStart || chunkStart > 1 ? { startNumber: chunkStart } : {}),
              listStyleType: style.listStyleType || undefined,
              items: pendingItems,
              x: base.x,
              y: pendingTop,
              width: base.width,
              height: Math.max(10, pendingBottom - pendingTop),
              style: extractTextStyle(style),
            })
            olRunningNumber += pendingItems.length
            pendingItems = []
            pendingTop = -1
            pendingBottom = -1
          }

          for (const li of liChildren) {
            const liImages = (Array.from(li.querySelectorAll('img')) as HTMLImageElement[]).filter(
              (img) => !isEmojiImg(img),
            )
            const liRect = li.getBoundingClientRect()
            const liY = liRect.top - slideRect.top
            const liBottom = liRect.bottom - slideRect.top

            // Extract inline badges from this <li> (same pattern as fast path)
            const { shapes: liSplitBadgeShapes, elements: liSplitBadgeEls, bgOnlyElements: liSplitBgOnlyEls } =
              extractInlineBadgeShapes(li, slideRect, liRect)
            const liLeadingOffset = computeLeadingOffset(
              liSplitBadgeShapes,
              liRect,
              slideRect,
            )
            const liSplitSkipBadges =
              liSplitBadgeEls.length > 0 ? new Set(liSplitBadgeEls) : false
            const liSplitStripBadges =
              liSplitBgOnlyEls.length > 0 ? new Set(liSplitBgOnlyEls) : false
            if (liSplitBadgeShapes.length > 0) elements.push(...liSplitBadgeShapes)

            if (liImages.length === 0) {
              // Check if this <li> has embedded code blocks
              const liCodeChildren = Array.from(li.children).filter((c) => {
                const t = c.tagName.toLowerCase()
                return (t === 'pre' || t === 'marp-pre') && !c.querySelector('svg')
              })
              // Normal <li>: accumulate into the running sub-list
              const liItems = extractListItemEl(
                li,
                0,
                liSplitSkipBadges,
                liSplitStripBadges,
                liLeadingOffset,
              )
              if (liCodeChildren.length === 0) {
                // No code blocks: just accumulate
                if (pendingTop < 0) pendingTop = liY
                pendingBottom = liBottom
                pendingItems.push(...liItems)
              } else {
                // Has code blocks: add text, flush, then emit code blocks
                if (pendingTop < 0) pendingTop = liY
                // Set pendingBottom to just before the first code block
                const firstCodeRect = liCodeChildren[0].getBoundingClientRect()
                pendingBottom = firstCodeRect.top - slideRect.top
                pendingItems.push(...liItems)
                flushPending()
                // Emit code block shapes
                for (const preChild of liCodeChildren) {
                  const preRect = preChild.getBoundingClientRect()
                  if (preRect.width <= 0 || preRect.height <= 0) continue
                  const preStyle = getComputedStyle(preChild)
                  const codeEl = preChild.querySelector('code')
                  const codeTarget = codeEl ?? preChild
                  const preBg = preStyle.backgroundColor
                  const codeBg = (preBg && preBg !== 'transparent' && preBg !== 'rgba(0, 0, 0, 0)' && preBg !== 'rgba(0,0,0,0)')
                    ? preBg
                    : (codeEl ? getComputedStyle(codeEl).backgroundColor : preBg)
                  const preCodeRuns = extractCodeRuns(codeTarget)
                  const preScaleFactor = computeAutoScaleFactor(
                    preCodeRuns,
                    preStyle,
                    preRect.height,
                    computeElementTransformScale(preChild),
                  )
                  const preBaseStyle = extractTextStyle(preStyle)
                  const preAdjustedFs = applyAutoScale(preCodeRuns, preBaseStyle.fontSize, preScaleFactor)
                  elements.push({
                    type: 'code',
                    text: codeTarget.textContent ?? '',
                    language: codeEl?.className?.replace('language-', '') ?? '',
                    runs: preCodeRuns,
                    x: preRect.left - slideRect.left,
                    y: preRect.top - slideRect.top,
                    width: preRect.width,
                    height: preRect.height,
                    style: {
                      ...preBaseStyle,
                      fontSize: preAdjustedFs,
                      backgroundColor: codeBg,
                    },
                  })
                }
              }
            } else {
              // <li> with embedded image(s):
              // 1. Include any text runs in this <li> in the pending sub-list
              const liItems = extractListItemEl(
                li,
                0,
                liSplitSkipBadges,
                liSplitStripBadges,
                liLeadingOffset,
              )
              if (liItems.length > 0) {
                if (pendingTop < 0) pendingTop = liY
                pendingBottom = liBottom
                pendingItems.push(...liItems)
              }
              // 2. Flush pending items (text before the image)
              flushPending()
              // 3. Emit the images at their actual rendered positions
              for (const img of liImages) {
                const imgRect = img.getBoundingClientRect()
                const imgFilter =
                  (getComputedStyle(img).filter && getComputedStyle(img).filter !== 'none')
                    ? getComputedStyle(img).filter
                    : undefined
                elements.push({
                  type: 'image',
                  src: img.src,
                  naturalWidth: img.naturalWidth,
                  naturalHeight: img.naturalHeight,
                  x: imgRect.left - slideRect.left,
                  y: imgRect.top - slideRect.top,
                  width: imgRect.width,
                  height: imgRect.height,
                  ...(imgFilter
                    ? { cssFilter: imgFilter, pageX: imgRect.left, pageY: imgRect.top }
                    : {}),
                })
              }
            }
          }

          flushPending()
        }
      } else if (tag === 'table') {
        const { rows: tableRows, colWidths } = extractTableData(child)
        elements.push({
          type: 'table',
          rows: tableRows,
          ...(colWidths.length > 0 ? { colWidths } : {}),
          ...base,
          style: extractTextStyle(style),
        })
        elements.push(...extractNestedImages(child, slideRect))
      } else if (tag === 'pre' || tag === 'marp-pre') {
        // If the <pre> contains a rendered SVG (e.g. Mermaid diagram), treat
        // it as an SVG image rather than a code block.
        const innerSvg = child.querySelector('svg')
        if (innerSvg) {
          try {
            const svgStr = new XMLSerializer().serializeToString(innerSvg)
            const b64 = btoa(unescape(encodeURIComponent(svgStr)))
            const dataUrl = `data:image/svg+xml;base64,${b64}`
            elements.push({
              type: 'image',
              src: dataUrl,
              naturalWidth: base.width,
              naturalHeight: base.height,
              ...base,
              // Request rasterization: Mermaid SVGs may use <foreignObject>
              // for text labels which PowerPoint cannot render from SVG data.
              // pageX/pageY are intentionally omitted: rasterizeSlideTargets
              // computes the absolute clip from the slide-relative x/y after
              // navigating to the correct slide (avoids stale bespoke-transform
              // coordinates).
              rasterize: true,
            })
          } catch {
            // Fall through to code block if SVG serialization fails
            const code = child.querySelector('code')
            const codeTarget = code ?? child
            // Resolve background: prefer <pre>, fall back to <code>
            const preBgCatch = style.backgroundColor
            const codeBgCatch = (preBgCatch && preBgCatch !== 'transparent' && preBgCatch !== 'rgba(0, 0, 0, 0)' && preBgCatch !== 'rgba(0,0,0,0)')
              ? preBgCatch
              : (code ? getComputedStyle(code).backgroundColor : preBgCatch)
            const codeRunsCatch = extractCodeRuns(codeTarget)
            const scaleFactorCatch = computeAutoScaleFactor(
              codeRunsCatch,
              style,
              rect.height,
              computeElementTransformScale(child),
            )
            const baseStyleCatch = extractTextStyle(style)
            const adjustedFsCatch = applyAutoScale(codeRunsCatch, baseStyleCatch.fontSize, scaleFactorCatch)
            elements.push({
              type: 'code',
              text: codeTarget.textContent ?? '',
              language: code?.className?.replace('language-', '') ?? '',
              runs: codeRunsCatch,
              ...base,
              style: {
                ...baseStyleCatch,
                fontSize: adjustedFsCatch,
                backgroundColor: codeBgCatch,
              },
            })
          }
        } else {
          const code = child.querySelector('code')
          const codeTarget = code ?? child
          // Resolve background: prefer <pre>, fall back to <code>
          const preBgStd = style.backgroundColor
          const codeBgStd = (preBgStd && preBgStd !== 'transparent' && preBgStd !== 'rgba(0, 0, 0, 0)' && preBgStd !== 'rgba(0,0,0,0)')
            ? preBgStd
            : (code ? getComputedStyle(code).backgroundColor : preBgStd)
          const codeRuns = extractCodeRuns(codeTarget)
          const scaleFactor = computeAutoScaleFactor(
            codeRuns,
            style,
            rect.height,
            computeElementTransformScale(child),
          )
          const baseStyle = extractTextStyle(style)
          const adjustedFontSize = applyAutoScale(codeRuns, baseStyle.fontSize, scaleFactor)
          elements.push({
            type: 'code',
            text: codeTarget.textContent ?? '',
            language: code?.className?.replace('language-', '') ?? '',
            runs: codeRuns,
            ...base,
            style: {
              ...baseStyle,
              fontSize: adjustedFontSize,
              backgroundColor: codeBgStd,
              ...(parseFloat(style.paddingTop) > 0
                ? { paddingTop: parseFloat(style.paddingTop) }
                : {}),
              ...(parseFloat(style.paddingRight) > 0
                ? { paddingRight: parseFloat(style.paddingRight) }
                : {}),
              ...(parseFloat(style.paddingBottom) > 0
                ? { paddingBottom: parseFloat(style.paddingBottom) }
                : {}),
              ...(parseFloat(style.paddingLeft) > 0
                ? { paddingLeft: parseFloat(style.paddingLeft) }
                : {}),
            },
          })
        }
      } else if (tag === 'img') {
        const img = child as HTMLImageElement
        // Emoji library images (Twemoji etc.) are captured as text by
        // extractTextRuns via the alt attribute — skip here to avoid
        // rendering them as a separate image element (which would duplicate).
        if (
          img.classList?.contains('emoji') ||
          img.src?.includes('twemoji') ||
          img.src?.includes('/emoji/')
        )
          return elements
        const imgFilter =
          style.filter && style.filter !== 'none' ? style.filter : undefined
        const clipped = resolveClippedRect(img, rect, slideRect)
        if (clipped.width <= 0 || clipped.height <= 0) return elements
        elements.push({
          type: 'image',
          src: img.src,
          naturalWidth: img.naturalWidth,
          naturalHeight: img.naturalHeight,
          x: clipped.left - slideRect.left,
          y: clipped.top - slideRect.top,
          width: clipped.width,
          height: clipped.height,
          // Store page-absolute coords when cssFilter is set so the export
          // tool can screenshot the rendered (filtered) region via Puppeteer.
          ...(clipped.isClipped
            ? { rasterize: true }
            : imgFilter
              ? { cssFilter: imgFilter, pageX: rect.left, pageY: rect.top }
              : {}),
        })
      } else if (tag === 'blockquote') {
        const borderWidth = parseFloat(style.borderLeftWidth) || 0
        const borderColor = style.borderLeftColor
        // Extract padding so slide-builder can inset the text box correctly.
        // paddingLeft provides the gap between the border-left bar and the text;
        // paddingTop/Right/Bottom are passed through for completeness.
        const paddingTop = parseFloat(style.paddingTop) || 0
        const paddingRight = parseFloat(style.paddingRight) || 0
        const paddingBottom = parseFloat(style.paddingBottom) || 0
        const paddingLeft = parseFloat(style.paddingLeft) || 0
        const { shapes: bqBadgeShapes, elements: bqBadgeEls, bgOnlyElements: bqBgOnlyEls } =
          extractInlineBadgeShapes(child, slideRect, rect)
        const bqLeadingOffset = computeLeadingOffset(
          bqBadgeShapes,
          rect,
          slideRect,
        )
        if (bqBadgeShapes.length > 0) elements.push(...bqBadgeShapes)
        const bqBadgeSet: boolean | Set<Element> =
          bqBadgeEls.length > 0 ? new Set(bqBadgeEls) : false
        const bqBgOnlySet: Set<Element> | false =
          bqBgOnlyEls.length > 0 ? new Set(bqBgOnlyEls) : false
        // Collect direct <ul>/<ol> children — these will be emitted as separate
        // ListElement objects so bullet markers and hierarchy are preserved.
        // They are excluded from extractTextRuns via skipBlockEls to prevent
        // the list text from being flattened into bullet-free plain text runs.
        const directListEls = Array.from(child.children).filter(c => {
          const t = c.tagName.toLowerCase()
          return t === 'ul' || t === 'ol'
        })
        const directListSet: Set<Element> | false =
          directListEls.length > 0 ? new Set<Element>(directListEls) : false
        elements.push({
          type: 'blockquote',
          runs: extractTextRuns(child, bqBadgeSet, bqBgOnlySet, directListSet),
          ...base,
          x: base.x + bqLeadingOffset,
          width: Math.max(10, base.width - bqLeadingOffset),
          style: {
            ...extractTextStyle(style),
            ...(paddingTop || paddingRight || paddingBottom || paddingLeft
              ? { paddingTop, paddingRight, paddingBottom, paddingLeft }
              : {}),
          },
          ...(borderWidth > 0
            ? { borderLeft: { width: borderWidth, color: borderColor } }
            : {}),
        })
        // Emit each direct list child as a proper ListElement so the PPTX
        // contains real bullet/numbering formatting rather than bare text.
        for (const listEl of directListEls) {
          const listRect = listEl.getBoundingClientRect()
          if (listRect.width <= 0 || listRect.height <= 0) continue
          const listStyle = getComputedStyle(listEl)
          const listTag = listEl.tagName.toLowerCase()
          const listOrdered = listTag === 'ol'
          const olStart = listOrdered ? ((listEl as HTMLOListElement).start || 1) : 1
          const olHasExplicitStart = listOrdered && listEl.hasAttribute('start')
          elements.push({
            type: 'list',
            ordered: listOrdered,
            ...(olHasExplicitStart ? { startNumber: olStart } : {}),
            listStyleType: listStyle.listStyleType || undefined,
            items: extractListItems(listEl, 0),
            x: listRect.left - slideRect.left,
            y: listRect.top - slideRect.top,
            width: listRect.width,
            height: listRect.height,
            style: extractTextStyle(listStyle),
          })
        }
        elements.push(...extractNestedImages(child, slideRect))
      } else if (tag === 'svg') {
        // Serialize SVG to a data URL for embedding as an image.
        // Use base64 encoding (not percent-encoding) because PptxGenJS and
        // Office require base64-encoded SVG in the PPTX XML <a:blip> element.
        // btoa(unescape(encodeURIComponent(...))) is the browser-safe way to
        // base64-encode a UTF-8 string without TextEncoder dependency.
        //
        // If the SVG contains <foreignObject> elements (e.g. mermaid@10 renders
        // flowchart text labels via foreignObject), PowerPoint cannot render
        // those elements natively.  Flag for rasterization so index.ts replaces
        // the SVG data URL with a PNG screenshot of the live browser rendering.
        try {
          const svgStr = new XMLSerializer().serializeToString(child)
          const b64 = btoa(unescape(encodeURIComponent(svgStr)))
          const dataUrl = `data:image/svg+xml;base64,${b64}`
          const hasForeignObject = child.querySelector('foreignObject') !== null
          const clipped = resolveClippedRect(child, rect, slideRect)
          if (clipped.width <= 0 || clipped.height <= 0) return elements
          elements.push({
            type: 'image',
            src: dataUrl,
            naturalWidth: clipped.width,
            naturalHeight: clipped.height,
            x: clipped.left - slideRect.left,
            y: clipped.top - slideRect.top,
            width: clipped.width,
            height: clipped.height,
            ...(hasForeignObject || clipped.isClipped
              ? { rasterize: true }
              : {}),
          })
        } catch {
          // Skip if serialization fails
        }
      } else if (tag === 'header' || tag === 'footer') {
        const { shapes: hfBadgeShapes, elements: hfBadgeEls, bgOnlyElements: hfBgOnlyEls } =
          extractInlineBadgeShapes(child, slideRect, rect)
        const hfLeadingOffset = computeLeadingOffset(
          hfBadgeShapes,
          rect,
          slideRect,
        )
        if (hfBadgeShapes.length > 0) elements.push(...hfBadgeShapes)
        const hfBadgeSet: boolean | Set<Element> =
          hfBadgeEls.length > 0 ? new Set(hfBadgeEls) : false
        const hfBgOnlySet: Set<Element> | false =
          hfBgOnlyEls.length > 0 ? new Set(hfBgOnlyEls) : false
        // Extend header/footer text box to the slide's right edge so PPTX
        // font metric differences (slightly wider glyphs) do not cause the
        // text to wrap when it fits on one line in the browser.
        // text-align is preserved, so right-aligned headers render correctly
        // even in a wider box.
        const hfWidth = Math.max(
          base.width - hfLeadingOffset,
          slideRect.width - base.x - hfLeadingOffset,
        )
        elements.push({
          type: tag,
          runs: extractTextRuns(child, hfBadgeSet, hfBgOnlySet),
          ...base,
          x: base.x + hfLeadingOffset,
          width: hfWidth,
          style: extractTextStyle(style),
        })
        elements.push(...extractNestedImages(child, slideRect))
      } else {
        const borders = extractCssBorders(style)
        const borderSides = [
          borders.top,
          borders.right,
          borders.bottom,
          borders.left,
        ].filter((border) => border !== undefined)
        const hasBorder = borderSides.length > 0
        const borderRadius = parseFloat(style.borderRadius) || 0
        const borderLeftWidth = borders.left?.width ?? 0
        const hasBorderLeft =
          borderLeftWidth > 0 && borderSides.length === 1
        const borderBottomWidth = borders.bottom?.width ?? 0
        const hasBorderBottom =
          borderBottomWidth > 0 && borderSides.length === 1
        // Detect visible box-shadow (used for card / elevated components)
        const boxShadow = style.boxShadow
        const hasBoxShadow = !!boxShadow && boxShadow !== 'none'
        const hasBackground =
          !!style.backgroundColor &&
          style.backgroundColor !== 'transparent' &&
          style.backgroundColor !== 'rgba(0, 0, 0, 0)' &&
          !style.backgroundColor.match(
            /rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*0(?:\.0+)?\s*\)/,
          )
        const blockChildren = walkElements(child, slideRect)

        const containerStyle = {
          backgroundColor: style.backgroundColor,
          ...(hasBorder
            ? {
                borders,
                borderWidth:
                  borders.top?.width ??
                  borders.right?.width ??
                  borders.bottom?.width ??
                  borders.left!.width,
                borderColor:
                  borders.top?.color ??
                  borders.right?.color ??
                  borders.bottom?.color ??
                  borders.left!.color,
                borderStyle:
                  borders.top?.style ??
                  borders.right?.style ??
                  borders.bottom?.style ??
                  borders.left?.style ??
                  'solid',
              }
            : {}),
          ...(borderRadius > 0 ? { borderRadius } : {}),
          ...(hasBorderLeft
            ? {
                borderLeft: {
                  width: borderLeftWidth,
                  color: borders.left!.color,
                },
              }
            : {}),
          ...(hasBorderBottom
            ? {
                borderBottom: {
                  width: borderBottomWidth,
                  color: borders.bottom!.color,
                  style: borders.bottom!.style,
                },
              }
            : {}),
          ...(hasBoxShadow ? { boxShadow: true } : {}),
        }

        const containerIsFlexOrGrid = /^(flex|inline-flex|grid|inline-grid)/.test(
          style.display,
        )
        const blockChildrenAllInlineLevel =
          !containerIsFlexOrGrid &&
          Array.from(child.children).every((node) => {
            const nodeStyle = getComputedStyle(node)
            return (
              nodeStyle.display === 'inline' ||
              /^inline-/.test(nodeStyle.display)
            )
          })

        if (blockChildren.length > 0) {
          const directElements = Array.from(child.children)
          const directOutputs = new Map<Element, SlideElement[]>()
          const unownedOutputs: SlideElement[] = []
          for (const output of blockChildren) {
            const owner = directElements.find((directElement) => {
              const directPath = sourceRefFor(directElement).domPath
              const outputPath = output.source?.domPath
              return (
                !!outputPath &&
                (outputPath === directPath ||
                  outputPath.startsWith(`${directPath}/`))
              )
            })
            if (owner) {
              const outputs = directOutputs.get(owner) ?? []
              outputs.push(output)
              directOutputs.set(owner, outputs)
            } else {
              unownedOutputs.push(output)
            }
          }
          const orderedChildren: SlideElement[] = []
          for (const node of Array.from(child.childNodes)) {
            if (node.nodeType === Node.ELEMENT_NODE) {
              const directTag = (node as Element).tagName.toLowerCase()
              if (directTag === 'svg' || directTag === 'mjx-container') {
                continue
              }
              const ownedOutputs = directOutputs.get(node as Element)
              if (ownedOutputs) {
                orderedChildren.push(...ownedOutputs)
              } else {
                const inlineSegments = extractInlineTextSegments(
                  node as Element,
                  slideRect,
                  new Set(),
                )
                if (inlineSegments.length > 0) {
                  orderedChildren.push(...inlineSegments)
                } else {
                  const inlineElement = node as Element
                  const inlineStyle = getComputedStyle(inlineElement)
                  if (inlineStyle.display === 'inline') {
                    const inlineRect = inlineElement.getBoundingClientRect()
                    const inlineRuns = extractTextRuns(inlineElement)
                    if (
                      inlineRuns.some(
                        (run) => !run.breakLine && run.text.trim() !== '',
                      )
                    ) {
                      orderedChildren.push({
                        type: 'paragraph',
                        source: sourceRefFor(inlineElement),
                        x: inlineRect.left - slideRect.left,
                        y: inlineRect.top - slideRect.top,
                        width: Math.max(10, inlineRect.width),
                        height: Math.max(10, inlineRect.height),
                        runs: inlineRuns,
                        style: extractTextStyle(inlineStyle),
                      })
                    }
                  }
                }
              }
            } else if (node.nodeType === Node.TEXT_NODE) {
              const text = node.textContent?.trim() ?? ''
              if (!text) continue
              const range = document.createRange()
              range.selectNodeContents(node)
              const hasMeasuredRange =
                typeof range.getBoundingClientRect === 'function'
              if (
                !hasMeasuredRange &&
                !containerIsFlexOrGrid &&
                !blockChildrenAllInlineLevel
              )
                continue
              const textRect =
                hasMeasuredRange
                  ? range.getBoundingClientRect()
                  : rect
              if (textRect.width <= 0 || textRect.height <= 0) continue
              const fallbackX = containerIsFlexOrGrid
                ? computeShallowFlexOffset(child, rect, slideRect, style)
                : 0
              orderedChildren.push({
                type: 'paragraph',
                source: {
                  ...sourceRefFor(child),
                  id: `${sourceRefFor(child).id}:text`,
                },
                x: hasMeasuredRange
                  ? textRect.left - slideRect.left
                  : base.x + fallbackX,
                y: textRect.top - slideRect.top,
                width: hasMeasuredRange
                  ? textRect.width
                  : Math.max(10, base.width - fallbackX),
                height: textRect.height,
                runs: [{ text }],
                style: extractTextStyle(style, false),
              })
            }
          }
          orderedChildren.push(...unownedOutputs)
          elements.push({
            type: 'container',
            children: orderedChildren,
            ...base,
            style: containerStyle,
          })
        } else {
          // Inline-only content (e.g. div with text, strong, br).
          // Emit a background/border box first (container with no children),
          // then the text as a paragraph element on top.
          if (hasBackground || hasBorder || hasBorderLeft || hasBorderBottom || hasBoxShadow) {
            elements.push({
              type: 'container',
              children: [],
              ...base,
              style: containerStyle,
            })
          }
          const runs = extractTextRuns(child)
          // Strip backgroundColor from runs that inherited the element's own
          // background-color.  The container shape drawn above already provides
          // the visual background; keeping the same colour as a per-run text
          // highlight causes visible colour bleed when text positioning drifts
          // slightly from the background shape.
          // Runs whose backgroundColor differs from the element background
          // (genuine inline highlights on a <span> or <mark>) are left untouched.
          if (hasBackground) {
            const elBg = style.backgroundColor
            for (const r of runs) {
              if (!r.breakLine && r.backgroundColor === elBg) {
                r.backgroundColor = undefined
              }
            }
          }
          if (runs.some((r) => !r.breakLine && r.text.trim() !== '')) {
            // When the element uses flexbox/grid vertical centering, emit
            // valign:'middle' so badge-style elements render correctly in PPTX.
            const valign: 'top' | 'middle' | 'bottom' =
              style.alignItems === 'center' ||
              style.justifyContent === 'center' ||
              style.verticalAlign === 'middle'
                ? 'middle'
                : 'top'
            // Extract CSS padding so the text box inset matches the div's
            // rendered padding.  For inline-only containers, both container and
            // paragraph share the same bounding rect (the div's outer rect); the
            // padding tells PptxGenJS how far to inset the text from the edges.
            const paddingTop = parseFloat(style.paddingTop) || 0
            const paddingRight = parseFloat(style.paddingRight) || 0
            const paddingBottom = parseFloat(style.paddingBottom) || 0
            const paddingLeft = parseFloat(style.paddingLeft) || 0
            // When the element is a direct flex/grid child and its runs include
            // emoji characters (converted from Twemoji <img> via alt text), the
            // bounding box width equals the intrinsic content width — fitted
            // exactly to the browser's rendering.  PowerPoint's Segoe UI Emoji
            // glyph may render slightly wider than the 1em Twemoji image, causing
            // the emoji to wrap to the next line.  Extend the text box to the
            // parent container's right edge to give extra room without
            // overlapping sibling flex items.
            // Cache parent's bounding rect once; used by emojiWidthOverride,
            // nowrapWidthOverride, and the parentIsFlexOrGrid width slack below.
            // getBoundingClientRect() triggers a forced reflow so a single call
            // per element is preferred.
            const parentRight = parentIsFlexOrGrid
              ? parent.getBoundingClientRect().right - slideRect.left
              : 0
            const emojiWidthOverride: number | undefined = (() => {
              if (!parentIsFlexOrGrid) return undefined
              const hasEmoji = runs.some(
                (r) =>
                  !r.breakLine &&
                  /\p{Extended_Pictographic}/u.test(r.text),
              )
              if (!hasEmoji) return undefined
              // Do not extend when a visible sibling element follows in DOM order
              // within the same flex/grid container.  That sibling occupies horizontal
              // space to the right; extending this element's text box to the parent
              // right edge would push a centred/aligned emoji into the sibling's area,
              // causing overlap.  (ADR-30)
              let sib = child.nextElementSibling
              while (sib !== null) {
                const sibStyle = getComputedStyle(sib as Element)
                if (sibStyle.display !== 'none' && sibStyle.visibility !== 'hidden') {
                  const sibRect = (sib as HTMLElement).getBoundingClientRect()
                  if (sibRect.width > 0 && sibRect.height > 0) return undefined
                }
                sib = sib.nextElementSibling
              }
              const extended = Math.max(base.width, parentRight - base.x)
              return extended > base.width ? extended : undefined
            })()
            // ADR-22: When a flex/grid child has white-space:nowrap, the
            // browser renders it on one line with a tight bounding box.
            // PPTX font metrics may be slightly wider, causing the text to
            // wrap.  Extend the text box width by 15% to absorb font-metric
            // variance.  The text box extends rightward (never leftward) to
            // avoid overlapping sibling flex items.  Since slide-builder sets
            // wrap:false for nowrap elements, text displays correctly even if
            // the box extends past the parent boundary.
            const nowrapWidthOverride: number | undefined = (() => {
              if (emojiWidthOverride !== undefined) return undefined
              if (!parentIsFlexOrGrid) return undefined
              if (style.whiteSpace !== 'nowrap') return undefined
              const wantedWidth = base.width * 1.15
              // Cap at slide right edge to avoid extending past the canvas
              const slideWidth = slideRect.width
              const maxWidth = Math.max(base.width, slideWidth - base.x)
              const finalWidth = Math.min(wantedWidth, maxWidth)
              return finalWidth > base.width ? finalWidth : undefined
            })()
            // Never shift x leftward — that causes overlap with sibling flex
            // items (slide 68 regression).  The text box extends rightward only.
            const nowrapXShift = 0
            elements.push({
              type: 'paragraph',
              runs,
              ...base,
              ...(emojiWidthOverride !== undefined
                ? { width: emojiWidthOverride }
                : nowrapWidthOverride !== undefined
                  ? { width: nowrapWidthOverride }
                  : // Inline-only containers (e.g. display:inline-block badges)
                    // have tight-fitting widths from browser font metrics.
                    // PowerPoint fonts may render slightly wider, causing text to
                    // wrap.  Add a small slack (8 px) when the container has a
                    // visible background (badge/chip pattern) to absorb the
                    // font-metric variance.
                    hasBackground
                    ? { width: base.width + 8 }
                    : parentIsFlexOrGrid
                      ? {
                          // ADR-23: Non-background text in flex/grid child containers
                          // (e.g. step-body divs, label divs) can wrap in PPTX when
                          // DirectWrite font metrics are slightly wider than Chrome's
                          // Skia. The wrapping causes the next absolutely-positioned
                          // item to overlap. Add 16px slack capped at parent right edge.
                          width: Math.min(
                            base.width + 16,
                            Math.max(base.width, parentRight - base.x),
                          ),
                        }
                      : {}),
              style: {
                ...extractTextStyle(style),
                ...(paddingTop || paddingRight || paddingBottom || paddingLeft
                  ? { paddingTop, paddingRight, paddingBottom, paddingLeft }
                  : {}),
              },
              valign,
            })
          }
          elements.push(...extractNestedImages(child, slideRect))
        }
      }

    return elements
  }

  function walkDisplayContents(
    container: Element,
    effectiveParent: Element,
    slideRect: DOMRect,
  ): SlideElement[] {
    const elements: SlideElement[] = []
    const containerStyle = getComputedStyle(container)
    const textIsVisible = containerStyle.visibility !== 'hidden'
    for (const node of Array.from(container.childNodes)) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const child = node as Element
        const walked = walkElement(effectiveParent, child, slideRect)
        if (walked.length > 0) {
          elements.push(...walked)
          continue
        }

        const childStyle = getComputedStyle(child)
        const childRect = child.getBoundingClientRect()
        if (
          childStyle.display !== 'inline' ||
          childStyle.visibility === 'hidden' ||
          childRect.width === 0 ||
          childRect.height === 0
        )
          continue
        const runs = extractTextRuns(child)
        if (runs.some((run) => !run.breakLine && run.text.trim() !== '')) {
          elements.push({
            type: 'paragraph',
            runs,
            x: childRect.left - slideRect.left,
            y: childRect.top - slideRect.top,
            width: childRect.width,
            height: childRect.height,
            style: extractTextStyle(childStyle),
          })
        }
        continue
      }
      if (node.nodeType !== Node.TEXT_NODE) continue
      if (!textIsVisible) continue

      const text = node.textContent ?? ''
      if (text.trim() === '') continue

      const range = document.createRange()
      range.selectNodeContents(node)
      const rect = range.getBoundingClientRect()
      if (rect.width === 0 || rect.height === 0) continue

      elements.push({
        type: 'paragraph',
        source: sourceRefFor(container),
        runs: [
          {
            text: text.trim(),
            color: containerStyle.color,
            fontSize: parseFloat(containerStyle.fontSize) || 16,
            fontFamily: containerStyle.fontFamily,
            bold: parseInt(containerStyle.fontWeight, 10) >= 600,
            italic: containerStyle.fontStyle === 'italic',
            underline:
              containerStyle.textDecorationLine?.includes('underline'),
            strikethrough:
              containerStyle.textDecorationLine?.includes('line-through'),
          },
        ],
        x: rect.left - slideRect.left,
        y: rect.top - slideRect.top,
        width: rect.width,
        height: rect.height,
        style: extractTextStyle(containerStyle),
      })
    }
    return elements
  }

  function walkElements(parent: Element, slideRect: DOMRect): SlideElement[] {
    const elements: SlideElement[] = []
    for (const child of Array.from(parent.children)) {
      elements.push(...walkElement(parent, child, slideRect))
    }
    return elements
  }

  function walkSectionNodes(
    parent: Element,
    slideRect: DOMRect,
  ): SlideElement[] {
    const elements: SlideElement[] = []
    const style = getComputedStyle(parent)
    if (style.display === 'none' || style.visibility === 'hidden') return elements
    for (const node of Array.from(parent.childNodes)) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const child = node as Element
        elements.push(...walkElement(parent, child, slideRect))
        continue
      }
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent?.trim() ?? ''
        if (!text) continue
        const range = document.createRange()
        range.selectNodeContents(node)
        if (typeof range.getBoundingClientRect !== 'function') continue
        const rect = range.getBoundingClientRect()
        if (rect.width <= 0 || rect.height <= 0) continue
        elements.push({
          type: 'paragraph',
          source: {
            ...sourceRefFor(parent),
            id: `${sourceRefFor(parent).id}:direct-text`,
          },
          x: rect.left - slideRect.left,
          y: rect.top - slideRect.top,
          width: rect.width,
          height: rect.height,
          runs: [{
            text,
            color: style.color,
            fontSize: parseFloat(style.fontSize) || 16,
            fontFamily: style.fontFamily,
            bold: parseInt(style.fontWeight, 10) >= 600,
            italic: style.fontStyle === 'italic',
            underline: style.textDecorationLine?.includes('underline'),
            strikethrough: style.textDecorationLine?.includes('line-through'),
          }],
          style: extractTextStyle(style, false),
        })
      }
    }
    return elements
  }

  // -----------------------------------------------------------------
  // Pseudo-element global-rule signatures.
  //
  // content:'' decorative pseudo-elements defined with a GLOBAL CSS rule
  // (e.g. `section::before { ... }` applied to all slides) share the same
  // background colour across both classless sections and classed sections.
  // We collect the background colours that appear on classless sections first,
  // then use them to suppress the same bar on classed sections — preventing
  // false banners when a global theme defines section::before for all slides
  // and some slides happen to carry user classes (cover, agenda, etc.).
  //
  // This is a mutable Set populated after slideGroups is built.
  // -----------------------------------------------------------------
  const globalPseudoSignatures = new Set<string>()

  // -----------------------------------------------------------------
  // Helper: extract visible ::before / ::after pseudo-elements as
  // coloured rectangle shapes.
  //
  // CSS pseudo-elements are not part of the DOM tree and cannot be
  // queried via querySelectorAll.  However, getComputedStyle(el, '::before')
  // returns their computed styles.  When a pseudo-element has a non-
  // transparent background-color and non-zero dimensions, we emit an editable
  // ContainerElement. Complex pseudo paint is captured as a raster image.
  // -----------------------------------------------------------------
  function extractPseudoElements(
    host: Element,
    slideRect: DOMRect,
    selectedPseudo?: '::before' | '::after',
  ): SlideElement[] {
    const shapes: SlideElement[] = []
    const pseudoElements = selectedPseudo
      ? [selectedPseudo]
      : (['::before', '::after'] as const)

    for (const pseudo of pseudoElements) {
      const ps = getComputedStyle(host, pseudo)
      // Pseudo-elements without content:'...' don't render, but
      // getComputedStyle still returns data.  Skip non-rendered pseudo-elements.
      // Empty-string pseudo-elements are valid CSS paint. Section-root cases
      // below retain the existing class/global-signature safeguards.
      const rawContent = ps.content
      if (!rawContent || rawContent === 'none' || rawContent === 'normal')
        continue
      const isSection = host.tagName.toLowerCase() === 'section'
      if (
        !isSection &&
        ps.position !== 'absolute' &&
        ps.position !== 'fixed'
      )
        continue
      // Check for empty-string content but do NOT skip — an empty string with a
      // visible background-color is a valid CSS decorative bar (e.g.
      // section::before { content: ''; background: blue; height: 12px; }).
      // HOWEVER: scoped per-slide pseudo-elements (from Marp's <style scoped>)
      // can leak across slides because Marp's CSS scoping uses shared
      // data-marpit-scope-* attributes.  These should only be extracted when
      // the section has a user-defined class (e.g. section.decorated) that
      // specifically triggers the rule.  Sections without a user class skip
      // content-empty pseudo-elements to avoid false banners.
      const stripped = rawContent.replace(/^["']|["']$/g, '').trim()
      if (stripped === '' && isSection) {
        const sectionClass = (host as HTMLElement).className?.trim() ?? ''
        if (!sectionClass) continue
        // Also skip when the same pseudo-element background appears on sections
        // without a user class — that means the CSS rule is global (e.g.
        // `section::before { background: navy }` for all slides) rather than
        // class-specific (e.g. `section.decorated::before`).  Extracting a
        // global bar only for classed slides creates inconsistent banners.
        const pgBg = ps.backgroundColor
        if (
          !pgBg ||
          pgBg === 'transparent' ||
          pgBg === 'rgba(0, 0, 0, 0)' ||
          globalPseudoSignatures.has(`${pseudo}:${pgBg}`)
        )
          continue
      }
      const bg = ps.backgroundColor
      const hasBackground =
        !!bg && bg !== 'transparent' && bg !== 'rgba(0, 0, 0, 0)'
      const hasBackgroundImage =
        !!ps.backgroundImage && ps.backgroundImage !== 'none'
      const pseudoBorders = {
        top: parseFloat(ps.borderTopWidth) || 0,
        right: parseFloat(ps.borderRightWidth) || 0,
        bottom: parseFloat(ps.borderBottomWidth) || 0,
        left: parseFloat(ps.borderLeftWidth) || 0,
      }
      const hasBorders = [
        pseudoBorders.top,
        pseudoBorders.right,
        pseudoBorders.bottom,
        pseudoBorders.left,
      ].some((width) => width > 0)
      const hasTransform = !!ps.transform && ps.transform !== 'none'
      const hasVisualPaint =
        hasBackground ||
        hasBackgroundImage ||
        hasBorders ||
        (!!ps.boxShadow && ps.boxShadow !== 'none') ||
        stripped !== ''
      if (!hasVisualPaint) continue

      const declaredWidth = parseFloat(ps.width) || 0
      const declaredHeight = parseFloat(ps.height) || 0
      const w = Math.max(
        declaredWidth,
        pseudoBorders.left + pseudoBorders.right,
      )
      const h = Math.max(
        declaredHeight,
        pseudoBorders.top + pseudoBorders.bottom,
      )
      if (w <= 0 || h <= 0) continue

      // Position: pseudo-elements with position:absolute or fixed use
      // top/left/right/bottom.  For bars that span the full width,
      // width is often 100% (resolved to the section width).
      const position = ps.position
      const hostRect = host.getBoundingClientRect()
      let x = hostRect.left - slideRect.left
      let y = hostRect.top - slideRect.top

      if (position === 'absolute' || position === 'fixed') {
        const top = parseFloat(ps.top)
        const left = parseFloat(ps.left)
        const bottom = parseFloat(ps.bottom)
        if (!isNaN(top)) y = hostRect.top - slideRect.top + top
        else if (!isNaN(bottom)) y = hostRect.top - slideRect.top + hostRect.height - bottom - h
        else y = hostRect.top - slideRect.top
        const right = parseFloat(ps.right)
        if (!isNaN(left)) x = hostRect.left - slideRect.left + left
        else if (!isNaN(right) && !isSection)
          x = hostRect.left - slideRect.left + hostRect.width - right - w
      }

      const clippedLeft = Math.max(0, x)
      const clippedTop = Math.max(0, y)
      const clippedRight = Math.min(slideRect.width, x + w)
      const clippedBottom = Math.min(slideRect.height, y + h)
      if (clippedRight <= clippedLeft || clippedBottom <= clippedTop) continue

      const base = {
        x: clippedLeft,
        y: clippedTop,
        width: clippedRight - clippedLeft,
        height: clippedBottom - clippedTop,
        ...extractStackingMetadata(ps),
        ...extractStackingMetadata(getComputedStyle(host)),
      }
      const source = sourceRefFor(host)
      const pseudoNeedsRaster =
        hasBackgroundImage ||
        hasBorders ||
        hasTransform ||
        (stripped !== '' && !isSection) ||
        (!!ps.boxShadow && ps.boxShadow !== 'none')

      shapes.push(
        pseudoNeedsRaster
          ? {
              type: 'image',
              source: {
                ...source,
                id: `${source.id}:${pseudo}`,
                domPath: `${source.domPath}/${pseudo}`,
              },
              src: '',
              naturalWidth: base.width,
              naturalHeight: base.height,
              ...base,
              rasterize: true,
            }
          : {
              type: 'container',
              source: {
                ...source,
                id: `${source.id}:${pseudo}`,
                domPath: `${source.domPath}/${pseudo}`,
              },
              children: [],
              ...base,
              style: { backgroundColor: bg },
            },
      )
    }

    return shapes
  }

  function extractDescendantPseudoElements(
    root: Element,
    slideRect: DOMRect,
  ): SlideElement[] {
    const elements: SlideElement[] = []
    for (const child of Array.from(root.querySelectorAll('*'))) {
      elements.push(...extractPseudoElements(child, slideRect, '::before'))
      elements.push(...extractPseudoElements(child, slideRect, '::after'))
    }
    return elements
  }

  function attachPseudoElements(
    elements: SlideElement[],
    pseudoElements: SlideElement[],
  ): SlideElement[] {
    const before = new Map<string, SlideElement[]>()
    const after = new Map<string, SlideElement[]>()
    const placed = new Set<SlideElement>()

    for (const pseudo of pseudoElements) {
      const sourceId = pseudo.source?.id
      if (!sourceId) continue
      const marker = sourceId.endsWith(':::after') ? 'after' : 'before'
      const hostId = sourceId.replace(/:::(before|after)$/, '')
      const map = marker === 'after' ? after : before
      const bucket = map.get(hostId) ?? []
      bucket.push(pseudo)
      map.set(hostId, bucket)
    }

    const nonContainerHosts = new Set<string>()
    const collectNonContainerHosts = (element: SlideElement) => {
      if (element.type !== 'container' && element.source?.id) {
        nonContainerHosts.add(element.source.id)
      }
      if (element.type === 'container') {
        element.children.forEach(collectNonContainerHosts)
      }
    }
    elements.forEach(collectNonContainerHosts)

    const visit = (element: SlideElement): SlideElement[] => {
      const sourceId = element.source?.id
      const canOwnPseudo =
        sourceId !== undefined &&
        (element.type !== 'container' || !nonContainerHosts.has(sourceId)) &&
        ![...(before.get(sourceId) ?? []), ...(after.get(sourceId) ?? [])]
          .some((pseudo) => placed.has(pseudo))
      const attachedBefore =
        canOwnPseudo ? (before.get(sourceId!) ?? []) : []
      const attachedAfter =
        canOwnPseudo ? (after.get(sourceId!) ?? []) : []
      attachedBefore.forEach((pseudo) => placed.add(pseudo))
      attachedAfter.forEach((pseudo) => placed.add(pseudo))

      if (element.type === 'container' && canOwnPseudo) {
        element = {
          ...element,
          children: [
            ...attachedBefore,
            ...element.children.flatMap(visit),
            ...attachedAfter,
          ],
        }
        return [element]
      }
      if (element.type === 'container') {
        element = {
          ...element,
          children: element.children.flatMap(visit),
        }
      }
      return [...attachedBefore, element, ...attachedAfter]
    }

    let attached = elements.flatMap(visit)
    const orphaned = pseudoElements.filter((pseudo) => !placed.has(pseudo))
    if (orphaned.length > 0) {
      const byParent = new Map<string, SlideElement[]>()
      for (const pseudo of orphaned) {
        if (!pseudo.source?.parentId) continue
        const bucket = byParent.get(pseudo.source.parentId) ?? []
        bucket.push(pseudo)
        byParent.set(pseudo.source.parentId, bucket)
      }
      const compareChildPaths = (
        left: SlideElement,
        right: SlideElement,
      ): number => {
        const leftPath = left.source?.domPath ?? ''
        const rightPath = right.source?.domPath ?? ''
        return leftPath.localeCompare(rightPath, undefined, { numeric: true })
      }
      const insert = (element: SlideElement): SlideElement => {
        if (element.type !== 'container') return element
        let children = element.children.map(insert)
        const additions = element.source?.id
          ? byParent.get(element.source.id)
          : undefined
        if (additions) {
          additions.forEach((pseudo) => placed.add(pseudo))
          children = [...children, ...additions].sort(compareChildPaths)
        }
        return { ...element, children }
      }
      attached = attached.map(insert)
    }
    return [
      ...attached,
      ...pseudoElements.filter((pseudo) => !placed.has(pseudo)),
    ]
  }

  // -----------------------------------------------------------------
  // Main logic — handle Marp Inline SVG mode with 3-layer sections
  //
  // When ![bg] is used, Marp generates 3 sections per slide:
  //   - data-marpit-advanced-background="background" (bg images)
  //   - data-marpit-advanced-background="content" (actual content)
  //   - data-marpit-advanced-background="pseudo" (page numbers etc.)
  // Without ![bg], there's just one section per slide.
  //
  // `data-marpit-pagination` is treated as a deck-wide source flag only.
  // The HTML pseudo-element text is not exported as ordinary text and we do
  // not reconstruct per-slide placement metadata here.
  // For `paginate: false`, Marp still emits one top-level <section> per slide
  // under <svg data-marpit-svg><foreignObject>..., but without the attribute.
  // -----------------------------------------------------------------
  const allSections = Array.from(root.querySelectorAll('section')).filter(
    (section) => {
      // Ignore nested sections inside slide content. Slide root sections are the
      // outermost section elements in the rendered Marp document.
      if (section.parentElement?.closest('section')) return false

      // Standard Marp output: section under svg > foreignObject
      if (section.parentElement?.tagName.toLowerCase() === 'foreignobject') {
        return true
      }

      // Test fixtures and simple DOMs may place the slide section directly.
      return section.hasAttribute('data-marpit-pagination')
    },
  )

  // Group sections by parent <svg> element.  Each Marp slide is wrapped
  // in its own <svg data-marpit-svg>, containing 1 section (no bg image)
  // or 3 sections (background/content/pseudo layers).  Using the SVG
  // index as grouping key is immune to:
  //   - paginate:hold collisions (same data-marpit-pagination on multiple slides)
  //   - paginate:false gaps (sections lacking data-marpit-pagination)
  const allSvgs = Array.from(
    document.querySelectorAll('svg[data-marpit-svg]'),
  )

  const slideGroups = new Map<
    string,
    { content?: Element; background?: Element; pseudo?: Element }
  >()

  for (const [index, section] of allSections.entries()) {
    // Navigate: section → foreignObject → svg
    const fo = section.parentElement
    const svg =
      fo?.tagName.toLowerCase() === 'foreignobject' ? fo.parentElement : null
    const key =
      svg?.hasAttribute('data-marpit-svg')
        ? String(allSvgs.indexOf(svg))
        : section.getAttribute('data-marpit-pagination') ??
          section.getAttribute('id') ??
          String(index)
    const layer = section.getAttribute('data-marpit-advanced-background')

    if (!slideGroups.has(key)) slideGroups.set(key, {})
    const entry = slideGroups.get(key)!

    if (layer === 'content') {
      entry.content = section
    } else if (layer === 'background') {
      entry.background = section
    } else if (layer === 'pseudo') {
      // Preserve the pseudo layer in the group so sourceHasPagination still
      // works even if pagination metadata lives only on that layer.
      entry.pseudo = section
    } else {
      // Non-inline-SVG mode or slide without ![bg]
      entry.content = section
    }
  }

  // Populate globalPseudoSignatures from classless content sections.
  // A background that appears on a classless section is a global theme rule
  // and must not be extracted as a banner for sections that do have a class.
  for (const { content } of slideGroups.values()) {
    const sec = content
    if (!sec) continue
    const secClass = (sec as HTMLElement).className?.trim() ?? ''
    if (secClass) continue // only scan classless sections
    for (const pseudo of ['::before', '::after'] as const) {
      const ps = getComputedStyle(sec, pseudo)
      const rawC = ps.content
      if (!rawC || rawC === 'none' || rawC === 'normal') continue
      if (rawC.replace(/^["']|["']$/g, '').trim() !== '') continue // non-empty content
      const bg = ps.backgroundColor
      if (!bg || bg === 'transparent' || bg === 'rgba(0, 0, 0, 0)') continue
      globalPseudoSignatures.add(`${pseudo}:${bg}`)
    }
  }

  return Array.from(slideGroups.values()).map(
    ({ content, background, pseudo }, slideIdx) => {
      const section = content ?? background!
      sourceRefFor(section)
      const sectionStyle = getComputedStyle(section)

      // For split-background slides (e.g. ![bg left:40%]), the content
      // section occupies only part of the SVG viewport.  We must use the
      // parent SVG's bounding rect as the coordinate reference frame so
      // that both text elements and background images are positioned in
      // the full 1280×720 PPTX slide space.
      let slideRect: DOMRect
      if (background && content) {
        const fo = section.parentElement
        const svg =
          fo?.tagName.toLowerCase() === 'foreignobject'
            ? fo.parentElement
            : null
        slideRect =
          svg?.hasAttribute('data-marpit-svg')
            ? svg.getBoundingClientRect()
            : section.getBoundingClientRect()
      } else {
        slideRect = section.getBoundingClientRect()
      }
      const sectionRect = slideRect

      // -----------------------------------------------------------------
      // Extract background images from ![bg] directive's background layer.
      // Each <figure> in the background layer corresponds to one background
      // image. Multiple figures occur for split layouts (e.g. ![bg left]).
      // -----------------------------------------------------------------
      const backgroundImages: BgImageData[] = []

      if (background) {
        const figures = background.querySelectorAll('figure')
        for (const fig of Array.from(figures)) {
          const figStyle = getComputedStyle(fig)
          if (!figStyle.backgroundImage || figStyle.backgroundImage === 'none')
            continue

          // Extract URL from background-image CSS value
          const urlMatch = figStyle.backgroundImage.match(
            /url\(["']?([^"')]+)["']?\)/,
          )
          if (!urlMatch) continue

          const figRect = fig.getBoundingClientRect()
          const cssFilter =
            figStyle.filter && figStyle.filter !== 'none'
              ? figStyle.filter
              : undefined

          // Detect background-size: contain (![bg fit]) — needs rasterization
          // to preserve aspect ratio letterboxing.
          const bgSize = figStyle.backgroundSize
          const isContain = bgSize === 'contain'

          backgroundImages.push({
            url: urlMatch[1],
            x: figRect.left - sectionRect.left,
            y: figRect.top - sectionRect.top,
            width: figRect.width || sectionRect.width,
            height: figRect.height || sectionRect.height,
            ...(cssFilter ? { cssFilter } : {}),
            ...(isContain ? { backgroundSizeContain: true } : {}),
            pageX: figRect.left,
            pageY: figRect.top,
          })
        }
      }

      // Fallback: section's own background-image (non-![bg] CSS background)
      // This is produced by `_backgroundImage` / `_backgroundColor` Marp
      // directives.  Mark with `fromCssFallback` so the export tool knows
      // it must rasterise the full slide to capture background-size/position
      // and any background-color overlay faithfully.
      if (backgroundImages.length === 0) {
        const bgImg = sectionStyle.backgroundImage
        if (bgImg && bgImg !== 'none') {
          const urlMatch = bgImg.match(/url\(["']?([^"')]+)["']?\)/)
          if (urlMatch) {
            backgroundImages.push({
              url: urlMatch[1],
              x: 0,
              y: 0,
              width: sectionRect.width,
              height: sectionRect.height,
              pageX: sectionRect.left,
              pageY: sectionRect.top,
              fromCssFallback: true,
            })
          } else if (/gradient\s*\(/.test(bgImg)) {
            // CSS gradient (linear-gradient, radial-gradient, etc.) — cannot
            // be reproduced natively in PPTX.  Mark for rasterization so the
            // export tool screenshots the rendered section background.
            backgroundImages.push({
              url: '', // placeholder — will be replaced by rasterized data URL
              x: 0,
              y: 0,
              width: sectionRect.width,
              height: sectionRect.height,
              pageX: sectionRect.left,
              pageY: sectionRect.top,
              fromCssFallback: true,
            })
          }
        }
      }

      return {
        width: sectionRect.width,
        height: sectionRect.height,
        background: findBackgroundColor(section),
        backgroundImages,
        sourceHasPagination: [content, background, pseudo].some((node) =>
          node?.hasAttribute('data-marpit-pagination'),
        ),
        elements: attachPseudoElements(
          [
            // Pseudo-element bars on the section itself retain slide z-order.
            ...extractPseudoElements(section, sectionRect),
            ...walkSectionNodes(section, sectionRect),
          ],
          extractDescendantPseudoElements(section, sectionRect),
        ),
        notes:
          (
            section.querySelector('[data-marpit-presenter-notes]') ??
            root.querySelector(`.bespoke-marp-note[data-index="${slideIdx}"]`)
          )?.textContent?.trim() ?? '',
      }
    },
  )
}
