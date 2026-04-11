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
} from './types'

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
  function extractTextStyle(style: CSSStyleDeclaration): TextStyle {
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
    }
  }

  function isSemanticInlineHighlightTag(tag: string): boolean {
    return tag === 'strong' || tag === 'mark' || tag === 'code'
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
  ): TextRun[] {
    const runs: TextRun[] = []

    const elementStyle = getComputedStyle(element)
    const elementBg = elementStyle.backgroundColor
    const elementHasBg =
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
          // Normal paragraph text in marp-core HTML has no \n text nodes
          // between inline elements, so this is safe for all element types.
          const newlineCount = (text.match(/\n/g) ?? []).length
          for (let i = 0; i < newlineCount; i++) {
            // Deduplicate: skip if the previous run is already a break.
            // Without this guard, a "<br>\n" sequence (br tag followed by
            // a whitespace text node) emits two consecutive breakLine runs,
            // which renders as a blank line in PPTX around emoji or block items.
            if (!lastIsBreak()) runs.push({ text: '', breakLine: true })
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
          const effectiveBg: string | undefined = (() => {
            if (hasBg && !alphaZero) return bg
            const bi = elStyle.backgroundImage
            if (!bi || bi === 'none' || !bi.includes('linear-gradient'))
              return undefined
            const colorMatches = bi.match(/rgba?\([^)]+\)/g)
            if (!colorMatches || colorMatches.length === 0) return undefined
            // Return the last color that is NOT transparent/fully-alpha
            for (let ci = colorMatches.length - 1; ci >= 0; ci--) {
              const c = colorMatches[ci]
              if (
                c !== 'rgba(0, 0, 0, 0)' &&
                c !== 'rgba(0,0,0,0)' &&
                !/rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*0\s*\)/.test(c)
              ) {
                return c
              }
            }
            return undefined
          })()
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
          const childRuns = extractTextRuns(el, skipBadges, stripBadges)
          // When this element has a background-only shape (in stripBadges), strip
          // backgroundColor from its runs — the shaped provides the visual bg.
          if (stripBadges instanceof Set && stripBadges.has(el)) {
            childRuns.forEach((r) => { if (!r.breakLine) r.backgroundColor = undefined })
          }
          runs.push(...childRuns)
        }
      }
    }

    if (runs.length > 0) {
      const combinedText = runs.map((r) => r.text).join('')
      items.push({
        text: combinedText.trim(),
        level,
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

    function walk(node: Node) {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? ''
        if (text === '') return
        const parent = node.parentElement ?? codeEl
        const style = getComputedStyle(parent)
        runs.push({
          text,
          color: style.color,
          fontSize:
            parseFloat(style.fontSize) ||
            parseFloat(defaultStyle.fontSize) ||
            16,
          fontFamily: style.fontFamily || defaultStyle.fontFamily,
          bold: parseInt(style.fontWeight, 10) >= 600,
          italic: style.fontStyle === 'italic',
        })
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
        if (isFirstRow) {
          // Capture rendered column widths from first row for proportional layout
          colWidths.push((td as HTMLElement).offsetWidth)
        }
        const cellBg = style.backgroundColor
        const cellHasBg =
          !!cellBg && cellBg !== 'transparent' && cellBg !== 'rgba(0, 0, 0, 0)'
        const effectiveBg = cellHasBg ? cellBg : trHasBg ? trBg : cellBg
        cells.push({
          text: td.textContent ?? '',
          runs: extractTextRuns(td),
          isHeader: td.tagName.toLowerCase() === 'th',
          style: {
            color: style.color,
            backgroundColor: effectiveBg,
            fontSize: parseFloat(style.fontSize) || 16,
            fontFamily: style.fontFamily,
            fontWeight: parseInt(style.fontWeight, 10) || 400,
            textAlign: style.textAlign || 'left',
            borderColor: style.borderColor,
          },
        })
      }

      rows.push({ cells })
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
      if (isInlineBadgeDisplay && containerRect && containerHasNonBadgeText) {
        const badgeSSLeft = iRect.left - slideRect.left
        if (badgeSSLeft > containerSSLeft + 8) {
          // Non-leading badge in mixed content: render a background-only shape
          // (correct border-radius, no text runs).  The badge text stays in the
          // parent paragraph run so text flow positioning is preserved;
          // extractTextRuns strips its backgroundColor so the round container
          // shape provides the visual background instead.
          badges.push({
            type: 'container',
            children: [],
            x: iRect.left - slideRect.left,
            y: iRect.top - slideRect.top,
            width: iRect.width,
            height: iRect.height,
            style: {
              backgroundColor: bg,
              ...(inlineBorderRadius > 0 ? { borderRadius: inlineBorderRadius } : {}),
            },
          })
          bgOnlyElements.push(el as Element)
          continue
        }
      }
      // extractTextRuns receives the badge elements in skipInlineBadges and
      // omits their text from the parent flow, preventing duplication.
      const br = inlineBorderRadius
      // Capture badge text so it can be rendered directly inside the shape.
      const badgeRuns = extractTextRuns(el as Element)
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

  // -----------------------------------------------------------------
  // Helper: collect IMG descendants of a container as ImageElements.
  // Called after processing text-bearing block elements (paragraph,
  // heading, blockquote, list, table) whose content is handled by
  // extractTextRuns / extractListItems, which skip <img> tags.
  // -----------------------------------------------------------------
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
      images.push({
        type: 'image',
        src: imgEl.src,
        naturalWidth: imgEl.naturalWidth,
        naturalHeight: imgEl.naturalHeight,
        x: rect.left - slideRect.left,
        y: rect.top - slideRect.top,
        width: rect.width,
        height: rect.height,
        ...(cssFilter ? { cssFilter, pageX: rect.left, pageY: rect.top } : {}),
      })
    }
    return images
  }

  // -----------------------------------------------------------------
  // Helper: walk child elements and classify them
  // -----------------------------------------------------------------
  function walkElements(parent: Element, slideRect: DOMRect): SlideElement[] {
    const elements: SlideElement[] = []

    for (const child of Array.from(parent.children)) {
      const style = getComputedStyle(child)
      if (style.display === 'none' || style.visibility === 'hidden') continue

      if ((child as HTMLElement).dataset?.marpitPresenterNotes !== undefined)
        continue

      // Skip Marp advanced background container (handled at slide level)
      if (
        (child as HTMLElement).dataset?.marpitAdvancedBackgroundContainer !==
        undefined
      )
        continue

      const rect = child.getBoundingClientRect()
      if (rect.width === 0 || rect.height === 0) continue

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
        continue

      const base = {
        x: rect.left - slideRect.left,
        y: rect.top - slideRect.top,
        width: rect.width,
        height: rect.height,
      }

      if (/^h[1-6]$/.test(tag)) {
        // Capture CSS border decorations (border-bottom for h1 rules, border-left
        // for h2 side bars). These are set by the theme's stylesheet and are not
        // visible in the DOM tree, only in computed styles.
        const borderBottomWidth = parseFloat(style.borderBottomWidth) || 0
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
        // Extract inline badge shapes. Shapes are always emitted so badges
        // render as rounded pill/circle elements in PPTX.  For leading badges
        // (at the paragraph's left edge), the paragraph text box is shifted
        // right to avoid overlap with the badge shape.
        const { shapes: paraBadgeShapes, elements: paraBadgeEls, bgOnlyElements: paraBgOnlyEls } =
          extractInlineBadgeShapes(child, slideRect, rect)
        const paraLeadingOffset = computeLeadingOffset(
          paraBadgeShapes,
          rect,
          slideRect,
        )
        if (paraBadgeShapes.length > 0) elements.push(...paraBadgeShapes)

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
          paraBgOnlyEls.length > 0 ? new Set(paraBgOnlyEls) : false
        const runs = extractTextRuns(child, paraBadgeSet, paraBgOnlySet)
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
          })
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

        if (!hasEmbeddedImage) {
          // Fast path: no embedded images.
          // Extract inline badge shapes from each <li> so that rounded-corner
          // badges (e.g. <span style="border-radius:8px;background:#c05621">)
          // are rendered as positioned shapes rather than flat text highlights,
          // matching the paragraph badge extraction behaviour.
          const liBadgeSets = new Map<Element, Set<Element>>()
          const liBgOnlySets = new Map<Element, Set<Element>>()
          const liLeadingOffsetMap = new Map<Element, number>()
          for (const li of liChildren) {
            const liRect = li.getBoundingClientRect()
            const { shapes: liBadgeShapes, elements: liBadgeEls, bgOnlyElements: liBgOnlyEls } =
              extractInlineBadgeShapes(li, slideRect, liRect)
            const liLeadingOffset = computeLeadingOffset(
              liBadgeShapes,
              liRect,
              slideRect,
            )
            if (liBadgeShapes.length > 0) {
              elements.push(...liBadgeShapes)
            }
            if (liBadgeEls.length > 0) {
              liBadgeSets.set(li, new Set(liBadgeEls))
            }
            if (liBgOnlyEls.length > 0) {
              liBgOnlySets.set(li, new Set(liBgOnlyEls))
            }
            if (liLeadingOffset > 0) {
              liLeadingOffsetMap.set(li, liLeadingOffset)
            }
          }
          elements.push({
            type: 'list',
            ordered: tag === 'ol',
            items: extractListItems(
              child,
              0,
              liBadgeSets.size > 0 ? liBadgeSets : undefined,
              liBgOnlySets.size > 0 ? liBgOnlySets : undefined,
              liLeadingOffsetMap.size > 0 ? liLeadingOffsetMap : undefined,
            ),
            ...base,
            style: extractTextStyle(style),
          })
          elements.push(...extractNestedImages(child, slideRect))
        } else {
          // Split the list around image-containing <li> elements so the
          // extracted images are interleaved at their actual positions.
          let pendingItems: ListItem[] = []
          let pendingTop = -1
          let pendingBottom = -1

          const flushPending = () => {
            if (pendingItems.length === 0) return
            elements.push({
              type: 'list',
              ordered: tag === 'ol',
              items: pendingItems,
              x: base.x,
              y: pendingTop,
              width: base.width,
              height: Math.max(10, pendingBottom - pendingTop),
              style: extractTextStyle(style),
            })
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
              // Normal <li>: accumulate into the running sub-list
              const liItems = extractListItemEl(
                li,
                0,
                liSplitSkipBadges,
                liSplitStripBadges,
                liLeadingOffset,
              )
              if (pendingTop < 0) pendingTop = liY
              pendingBottom = liBottom
              pendingItems.push(...liItems)
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
      } else if (tag === 'pre') {
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
            elements.push({
              type: 'code',
              text: codeTarget.textContent ?? '',
              language: code?.className?.replace('language-', '') ?? '',
              runs: extractCodeRuns(codeTarget),
              ...base,
              style: {
                ...extractTextStyle(style),
                backgroundColor: style.backgroundColor,
              },
            })
          }
        } else {
          const code = child.querySelector('code')
          const codeTarget = code ?? child
          elements.push({
            type: 'code',
            text: codeTarget.textContent ?? '',
            language: code?.className?.replace('language-', '') ?? '',
            runs: extractCodeRuns(codeTarget),
            ...base,
            style: {
              ...extractTextStyle(style),
              backgroundColor: style.backgroundColor,
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
          continue
        const imgFilter =
          style.filter && style.filter !== 'none' ? style.filter : undefined
        elements.push({
          type: 'image',
          src: img.src,
          naturalWidth: img.naturalWidth,
          naturalHeight: img.naturalHeight,
          ...base,
          // Store page-absolute coords when cssFilter is set so the export
          // tool can screenshot the rendered (filtered) region via Puppeteer.
          ...(imgFilter
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
        elements.push({
          type: 'blockquote',
          runs: extractTextRuns(child, bqBadgeSet, bqBgOnlySet),
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
          elements.push({
            type: 'image',
            src: dataUrl,
            naturalWidth: base.width,
            naturalHeight: base.height,
            ...base,
            ...(hasForeignObject ? { rasterize: true } : {}),
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
        const borderTopWidth = parseFloat(style.borderTopWidth) || 0
        const borderTopStyle = style.borderTopStyle
        const hasBorder = borderTopWidth > 0 && borderTopStyle !== 'none'
        const borderRadius = parseFloat(style.borderRadius) || 0
        // Detect CSS border-left (used for note-box bar decorations)
        const borderLeftWidth = parseFloat(style.borderLeftWidth) || 0
        const borderLeftStyle = style.borderLeftStyle
        const hasBorderLeft =
          borderLeftWidth > 0 && borderLeftStyle !== 'none' && !hasBorder
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
            ? { borderWidth: borderTopWidth, borderColor: style.borderTopColor, borderStyle: borderTopStyle }
            : {}),
          ...(borderRadius > 0 ? { borderRadius } : {}),
          ...(hasBorderLeft
            ? {
                borderLeft: {
                  width: borderLeftWidth,
                  color: style.borderLeftColor,
                },
              }
            : {}),
          ...(hasBoxShadow ? { boxShadow: true } : {}),
        }

        const containerIsFlexOrGrid = /^(flex|inline-flex|grid|inline-grid)/.test(
          style.display,
        )

        if (blockChildren.length > 0) {
          // Block-level children → normal container element
          elements.push({
            type: 'container',
            children: blockChildren,
            ...base,
            style: containerStyle,
          })
          // A flex/grid container may have BOTH block-level child elements
          // (already captured in blockChildren) AND direct text nodes or
          // inline-only elements that are not themselves block-level.
          // These direct text nodes are not walked by walkElements (which
          // only iterates element.children, skipping Text nodes) and are
          // therefore silently dropped unless we handle them here.
          //
          // Example: <div style="display:flex">
          //            <span class="badge">1</span>   ← block child
          //            Agenda item text                ← direct text node
          //          </div>
          //
          // We perform a shallow pass over the container's childNodes.
          // For flex/grid containers, direct element children are already
          // emitted as blockChildren, so we only need to recover TEXT_NODEs.
          // For normal block containers, direct inline children are skipped by
          // walkElements and must be recovered here together with TEXT_NODEs.
          const shallowRuns: TextRun[] = []
          for (const node of Array.from(child.childNodes)) {
            if (node.nodeType === Node.TEXT_NODE) {
              // Only recover direct text nodes for flex/grid containers.
              // In block containers, direct text nodes alongside block children
              // are typically invisible source code — for example, mermaid.js
              // diagram source that the library replaces with an SVG element.
              // If the CDN script hasn't fully finished at DOM-walk time, the
              // orphaned text node (raw diagram syntax) would otherwise appear
              // as a text box rendered on top of the SVG image in the PPTX.
              if (!containerIsFlexOrGrid) continue
              const text = (node.textContent ?? '').trim()
              if (text !== '') {
                const childStyle = getComputedStyle(child)
                shallowRuns.push({
                  text,
                  color: childStyle.color,
                  fontSize: parseFloat(childStyle.fontSize) || 16,
                  fontFamily: childStyle.fontFamily,
                  bold: parseInt(childStyle.fontWeight, 10) >= 600,
                  italic: childStyle.fontStyle === 'italic',
                  underline: childStyle.textDecorationLine?.includes('underline'),
                  strikethrough: childStyle.textDecorationLine?.includes('line-through'),
                })
              }
            } else if (node.nodeType === Node.ELEMENT_NODE) {
              if (containerIsFlexOrGrid) continue

              const nodeEl = node as Element
              const nodeTag = nodeEl.tagName.toLowerCase()
              // SVG elements are always captured as images by walkElements.
              // Calling extractTextRuns on them would extract diagram label
              // text (e.g., mermaid node labels) as spurious prose runs.
              if (nodeTag === 'svg') continue
              const nodeStyle = getComputedStyle(nodeEl)
              // Only capture strictly inline elements (display === 'inline').
              // Elements with display: inline-block / inline-flex / inline-grid
              // are already processed by walkElements (its skip filter only
              // excludes display === 'inline') and they appear in blockChildren.
              // Including them here would duplicate their text as a sibling
              // paragraph — e.g. a standalone inline-block badge inside a div
              // would emit both its own container+paragraph AND an extra
              // paragraph from the shallow walk (slide 48 regression).
              if (nodeStyle.display === 'inline') {
                const inlineRuns = extractTextRuns(nodeEl)
                shallowRuns.push(...inlineRuns)
              }
            }
          }
          // Trim leading/trailing breaks
          while (
            shallowRuns.length > 0 &&
            shallowRuns[shallowRuns.length - 1].breakLine
          )
            shallowRuns.pop()
          while (shallowRuns.length > 0 && shallowRuns[0].breakLine)
            shallowRuns.shift()
          if (shallowRuns.some((r) => !r.breakLine && r.text.trim() !== '')) {
            const shallowLeadingOffset = containerIsFlexOrGrid
              ? computeShallowFlexOffset(child, rect, slideRect, style)
              : 0
            elements.push({
              type: 'paragraph',
              runs: shallowRuns,
              ...base,
              x: base.x + shallowLeadingOffset,
              width: Math.max(10, base.width - shallowLeadingOffset),
              style: extractTextStyle(style),
            })
          }
        } else {
          // Inline-only content (e.g. div with text, strong, br).
          // Emit a background/border box first (container with no children),
          // then the text as a paragraph element on top.
          if (hasBackground || hasBorder || hasBorderLeft || hasBoxShadow) {
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
            const emojiWidthOverride: number | undefined = (() => {
              if (!parentIsFlexOrGrid) return undefined
              const hasEmoji = runs.some(
                (r) =>
                  !r.breakLine &&
                  /\p{Extended_Pictographic}/u.test(r.text),
              )
              if (!hasEmoji) return undefined
              const parentRight =
                parent.getBoundingClientRect().right - slideRect.left
              const extended = Math.max(base.width, parentRight - base.x)
              return extended > base.width ? extended : undefined
            })()
            elements.push({
              type: 'paragraph',
              runs,
              ...base,
              ...(emojiWidthOverride !== undefined
                ? { width: emojiWidthOverride }
                : // Inline-only containers (e.g. display:inline-block badges)
                  // have tight-fitting widths from browser font metrics.
                  // PowerPoint fonts may render slightly wider, causing text to
                  // wrap.  Add a small slack (8 px) when the container has a
                  // visible background (badge/chip pattern) to absorb the
                  // font-metric variance.
                  hasBackground
                  ? { width: base.width + 8 }
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
  // transparent background-color and non-zero dimensions, we emit a
  // ContainerElement so it appears as a filled rectangle in the PPTX.
  // -----------------------------------------------------------------
  function extractPseudoElements(
    section: Element,
    slideRect: DOMRect,
  ): SlideElement[] {
    const shapes: SlideElement[] = []

    for (const pseudo of ['::before', '::after'] as const) {
      const ps = getComputedStyle(section, pseudo)
      // Pseudo-elements without content:'...' don't render, but
      // getComputedStyle still returns data.  Skip non-rendered pseudo-elements.
      // Also skip content:'""' — empty-string pseudo-elements (common in theme
      // CSS for decorative bars via background-color) should NOT be extracted
      // because they appear in HTML/preview but the user has not explicitly placed
      // them as content.  Extracting them creates phantom banners in PPTX that
      // don't match the user's intent.
      const rawContent = ps.content
      if (!rawContent || rawContent === 'none' || rawContent === 'normal')
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
      if (stripped === '') {
        const sectionClass = (section as HTMLElement).className?.trim() ?? ''
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
      if (!bg || bg === 'transparent' || bg === 'rgba(0, 0, 0, 0)') continue
      // Parse dimensions — pseudo-elements use width/height from CSS
      const w = parseFloat(ps.width) || 0
      const h = parseFloat(ps.height) || 0
      if (w === 0 && h === 0) continue

      // Position: pseudo-elements with position:absolute or fixed use
      // top/left/right/bottom.  For bars that span the full width,
      // width is often 100% (resolved to the section width).
      const position = ps.position
      let x = 0
      let y = 0

      if (position === 'absolute' || position === 'fixed') {
        const top = parseFloat(ps.top)
        const left = parseFloat(ps.left)
        const bottom = parseFloat(ps.bottom)
        if (!isNaN(top)) y = top
        else if (!isNaN(bottom)) y = slideRect.height - bottom - h
        if (!isNaN(left)) x = left
      }

      const effectiveW = w || slideRect.width
      const effectiveH = h || 0

      if (effectiveH <= 0) continue

      shapes.push({
        type: 'container',
        children: [],
        x,
        y,
        width: effectiveW,
        height: effectiveH,
        style: {
          backgroundColor: bg,
        },
      })
    }

    return shapes
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

  // Group sections by pagination number, tracking layers
  const slideGroups = new Map<
    string,
    { content?: Element; background?: Element; pseudo?: Element }
  >()

  for (const [index, section] of allSections.entries()) {
    const key =
      section.getAttribute('data-marpit-pagination') ??
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
      const sectionRect = section.getBoundingClientRect()
      const sectionStyle = getComputedStyle(section)

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

          backgroundImages.push({
            url: urlMatch[1],
            x: figRect.left - sectionRect.left,
            y: figRect.top - sectionRect.top,
            width: figRect.width || sectionRect.width,
            height: figRect.height || sectionRect.height,
            ...(cssFilter ? { cssFilter } : {}),
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
        elements: [
          // Pseudo-element bars (::before/::after) go behind content
          ...extractPseudoElements(section, sectionRect),
          ...walkElements(section, sectionRect),
        ],
        notes:
          (
            section.querySelector('[data-marpit-presenter-notes]') ??
            root.querySelector(`.bespoke-marp-note[data-index="${slideIdx}"]`)
          )?.textContent?.trim() ?? '',
      }
    },
  )
}
