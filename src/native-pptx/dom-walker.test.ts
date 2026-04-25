import { JSDOM } from 'jsdom'
import { extractSlides } from './dom-walker'
import { toListTextProps } from './slide-builder'

// ---------------------------------------------------------------------------
// Manual JSDOM setup (jest-environment-jsdom hangs with jest 30 + node 22)
// ---------------------------------------------------------------------------

let dom: InstanceType<typeof JSDOM>

beforeEach(() => {
  dom = new JSDOM('<!DOCTYPE html><html><body></body></html>')

  // Expose browser globals that dom-walker.ts relies on
  ;(globalThis as any).document = dom.window.document
  ;(globalThis as any).getComputedStyle = dom.window.getComputedStyle.bind(
    dom.window,
  )
  ;(globalThis as any).Node = dom.window.Node
  ;(globalThis as any).NodeFilter = dom.window.NodeFilter
  ;(globalThis as any).XMLSerializer = dom.window.XMLSerializer
})

afterEach(() => {
  delete (globalThis as any).document
  delete (globalThis as any).getComputedStyle
  delete (globalThis as any).Node
  delete (globalThis as any).NodeFilter
  delete (globalThis as any).XMLSerializer
  dom.window.close()
})

// ---------------------------------------------------------------------------
// Mock helpers (jsdom does not compute layout)
// ---------------------------------------------------------------------------

function mockRect(
  el: Element,
  rect: { left: number; top: number; width: number; height: number },
) {
  ;(el as any).getBoundingClientRect = () => ({
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
    right: rect.left + rect.width,
    bottom: rect.top + rect.height,
    x: rect.left,
    y: rect.top,
    // eslint-disable-next-line @typescript-eslint/no-empty-function
    toJSON() {},
  })
}

const defaultStyles: Record<string, string> = {
  display: 'block',
  visibility: 'visible',
  color: 'rgb(0, 0, 0)',
  backgroundColor: 'rgba(0, 0, 0, 0)',
  backgroundImage: 'none',
  fontSize: '16px',
  fontFamily: 'Arial',
  fontWeight: '400',
  fontStyle: 'normal',
  textAlign: 'left',
  textDecorationLine: 'none',
  lineHeight: '24px',
  borderColor: 'rgb(0, 0, 0)',
}

/**
 * Patch `getComputedStyle` so that it returns mocked style objects for the
 * specified elements. Optional pseudo-element mappings can be provided for
 * tests that need resolved `::before` / `::after` styles. Other elements fall
 * through to the original implementation.
 */
function mockStyles(
  mappings: [Element, Record<string, string>][],
  pseudoMappings: [Element, '::before' | '::after', Record<string, string>][] = [],
): () => void {
  const original = globalThis.getComputedStyle
  const map = new Map<Element, CSSStyleDeclaration>()
  const pseudoMap = new Map<Element, Map<string, CSSStyleDeclaration>>()

  function createStyleProxy(styles: Record<string, string>): CSSStyleDeclaration {
    const merged = { ...defaultStyles, ...styles }
    return new Proxy({} as CSSStyleDeclaration, {
      get(_t, prop: string) {
        if (prop === 'getPropertyValue')
          return (name: string) => merged[name] ?? ''
        return merged[prop] ?? ''
      },
    })
  }

  for (const [el, styles] of mappings) {
    map.set(el, createStyleProxy(styles))
  }

  for (const [el, pseudo, styles] of pseudoMappings) {
    if (!pseudoMap.has(el)) pseudoMap.set(el, new Map())
    pseudoMap.get(el)!.set(pseudo, createStyleProxy(styles))
  }

  ;(globalThis as any).getComputedStyle = (
    target: Element,
    pseudoElement?: string,
  ) => {
    if (pseudoElement) {
      return pseudoMap.get(target)?.get(pseudoElement)
        ?? original(target, pseudoElement)
    }
    return map.get(target) ?? original(target)
  }

  return () => {
    ;(globalThis as any).getComputedStyle = original
  }
}

/**
 * Set up a single-slide HTML document and mock the section element.
 */
function setupSlide(
  bodyContent: string,
  sectionRect = { left: 0, top: 0, width: 1280, height: 720 },
) {
  // Wrap in SVG > foreignObject so the section is recognised as a Marp slide
  // via the parentElement check (no data-marpit-pagination needed).
  // This keeps fixtures minimal while still exercising the normal Marp slide
  // extraction path used for bespoke HTML output.
  document.body.innerHTML = `
    <div id=":$p">
      <svg data-marpit-svg="" viewBox="0 0 1280 720">
        <foreignObject width="1280" height="720">
          <section>${bodyContent}</section>
        </foreignObject>
      </svg>
    </div>
  `
  const section = document.querySelector('section')!
  mockRect(section, sectionRect)
  return { section }
}

// -----------------------------------------------------------------------
// extractSlides — basic
// -----------------------------------------------------------------------

describe('extractSlides', () => {
  it('returns empty array for empty document', () => {
    expect(extractSlides()).toEqual([])
  })

  it('extracts top-level section as slide even without pagination', () => {
    document.body.innerHTML = `
      <div id=":$p">
        <svg data-marpit-svg="" viewBox="0 0 1280 720">
          <foreignObject width="1280" height="720">
            <section id="1" data-theme="default" lang="ja-JP">
              <h1>Title</h1>
            </section>
          </foreignObject>
        </svg>
      </div>
    `

    const section = document.querySelector('section')!
    const h1 = section.querySelector('h1')!

    mockRect(section, { left: 0, top: 0, width: 1280, height: 720 })
    mockRect(h1, { left: 70, top: 80, width: 1140, height: 60 })

    const restore = mockStyles([
      [
        section,
        { backgroundColor: 'rgb(255, 255, 255)', backgroundImage: 'none' },
      ],
      [h1, { fontSize: '40px', fontWeight: '700', color: 'rgb(34, 68, 102)' }],
    ])

    const slides = extractSlides()
    expect(slides).toHaveLength(1)
    expect(slides[0].elements).toHaveLength(1)
    expect(slides[0].elements[0].type).toBe('heading')

    restore()
  })

  it('extracts slide section (foreignObject) as one slide with correct dimensions', () => {
    const { section } = setupSlide('<h1>Title</h1>')
    const h1 = section.querySelector('h1')!

    mockRect(h1, { left: 70, top: 80, width: 1140, height: 60 })

    const restore = mockStyles([
      [
        section,
        { backgroundColor: 'rgb(255, 255, 255)', backgroundImage: 'none' },
      ],
      [h1, { fontSize: '40px', fontWeight: '700', color: 'rgb(34, 68, 102)' }],
    ])

    const slides = extractSlides()
    expect(slides).toHaveLength(1)
    expect(slides[0].width).toBe(1280)
    expect(slides[0].height).toBe(720)
    expect(slides[0].background).toBe('rgb(255, 255, 255)')
    expect(slides[0].elements).toHaveLength(1)
    expect(slides[0].elements[0].type).toBe('heading')

    restore()
  })

  it('extracts presenter notes', () => {
    const { section } = setupSlide(
      '<div data-marpit-presenter-notes>Speaker notes here</div>',
    )

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255, 255, 255)' }],
    ])

    const slides = extractSlides()
    expect(slides[0].notes).toBe('Speaker notes here')

    restore()
  })

  it('extracts presenter notes from bespoke HTML format (.bespoke-marp-note)', () => {
    // marp-cli bespoke output does NOT put notes in [data-marpit-presenter-notes].
    // Instead it injects <div class="bespoke-marp-note" data-index="N"> elements
    // as siblings to the slide section.
    const { section } = setupSlide('')
    const noteEl = document.createElement('div')
    noteEl.className = 'bespoke-marp-note'
    noteEl.setAttribute('data-index', '0')
    noteEl.innerHTML = '<p>Bespoke note line 1</p><p>line 2</p>'
    document.body.appendChild(noteEl)

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255, 255, 255)' }],
    ])

    const slides = extractSlides()
    expect(slides[0].notes).toBe('Bespoke note line 1line 2')

    restore()
  })
})

// -----------------------------------------------------------------------
// walkElements — tested through extractSlides
// -----------------------------------------------------------------------

describe('walkElements (via extractSlides)', () => {
  it('classifies heading elements as heading', () => {
    const { section } = setupSlide('<h1>Title</h1>')
    const h1 = section.querySelector('h1')!

    mockRect(h1, { left: 10, top: 20, width: 500, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [h1, { fontSize: '32px', fontWeight: '700', color: 'rgb(0, 0, 0)' }],
    ])

    const slides = extractSlides()
    const elements = slides[0].elements
    expect(elements).toHaveLength(1)
    expect(elements[0]).toMatchObject({
      type: 'heading',
      level: 1,
      x: 10,
      y: 20,
      width: 500,
      height: 40,
    })

    restore()
  })

  it('classifies paragraph elements as paragraph', () => {
    const { section } = setupSlide('<p>Hello world</p>')
    const p = section.querySelector('p')!

    mockRect(p, { left: 10, top: 50, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { color: 'rgb(51, 51, 51)' }],
    ])

    const slides = extractSlides()
    expect(slides[0].elements).toHaveLength(1)
    expect(slides[0].elements[0].type).toBe('paragraph')

    restore()
  })

  it('classifies list elements as list', () => {
    const { section } = setupSlide('<ul><li>Item 1</li><li>Item 2</li></ul>')
    const ul = section.querySelector('ul')!
    const lis = section.querySelectorAll('li')

    mockRect(ul, { left: 10, top: 100, width: 600, height: 48 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ul, {}],
      [lis[0], {}],
      [lis[1], {}],
    ])

    const slides = extractSlides()
    expect(slides[0].elements).toHaveLength(1)
    expect(slides[0].elements[0]).toMatchObject({
      type: 'list',
      ordered: false,
    })

    const listEl = slides[0].elements[0] as any
    expect(listEl.items).toHaveLength(2)

    restore()
  })

  it('skips hidden elements', () => {
    const { section } = setupSlide(
      '<p id="visible">Visible</p><p id="hidden">Hidden</p>',
    )
    const visible = document.getElementById('visible')!
    const hidden = document.getElementById('hidden')!

    mockRect(visible, { left: 0, top: 0, width: 600, height: 24 })
    mockRect(hidden, { left: 0, top: 30, width: 600, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [visible, {}],
      [hidden, { display: 'none' }],
    ])

    const slides = extractSlides()
    expect(slides[0].elements).toHaveLength(1)

    restore()
  })

  it('skips zero-size elements', () => {
    const { section } = setupSlide('<p>Empty</p>')
    const p = section.querySelector('p')!

    mockRect(p, { left: 0, top: 0, width: 0, height: 0 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {}],
    ])

    const slides = extractSlides()
    expect(slides[0].elements).toHaveLength(0)

    restore()
  })

  it('recursively expands div containers', () => {
    const { section } = setupSlide(
      '<div id="container"><p>Nested paragraph</p></div>',
    )
    const container = document.getElementById('container')!
    const p = container.querySelector('p')!

    mockRect(container, { left: 0, top: 0, width: 640, height: 200 })
    mockRect(p, { left: 10, top: 10, width: 620, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [container, { backgroundColor: 'rgb(200, 200, 200)' }],
      [p, {}],
    ])

    const slides = extractSlides()
    expect(slides[0].elements).toHaveLength(1)
    expect(slides[0].elements[0]).toMatchObject({ type: 'container' })
    expect((slides[0].elements[0] as any).children).toHaveLength(1)
    expect((slides[0].elements[0] as any).children[0].type).toBe('paragraph')

    restore()
  })

  it('computes slide-relative coordinates', () => {
    const { section } = setupSlide('<h2>Sub</h2>', {
      left: 100,
      top: 200,
      width: 1280,
      height: 720,
    })
    const h2 = section.querySelector('h2')!

    mockRect(h2, { left: 170, top: 280, width: 500, height: 36 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [h2, { fontSize: '28px', fontWeight: '700' }],
    ])

    const slides = extractSlides()
    expect(slides[0].elements[0]).toMatchObject({ x: 70, y: 80 })

    restore()
  })
})

// -----------------------------------------------------------------------
// extractTextRuns — tested through extractSlides
// -----------------------------------------------------------------------

describe('extractTextRuns (via extractSlides)', () => {
  it('extracts plain text runs', () => {
    const { section } = setupSlide('<p id="t">Hello world</p>')
    const p = document.getElementById('t')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        p,
        {
          color: 'rgb(0, 0, 0)',
          fontSize: '16px',
          fontWeight: '400',
          fontStyle: 'normal',
          textDecorationLine: 'none',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.runs).toHaveLength(1)
    expect(el.runs[0].text).toBe('Hello world')
    expect(el.runs[0].bold).toBe(false)

    restore()
  })

  it('parses bold and italic inline elements', () => {
    const { section } = setupSlide(
      '<p id="t">Normal <strong>Bold</strong> <em>Italic</em></p>',
    )
    const p = document.getElementById('t')!
    const strong = p.querySelector('strong')!
    const em = p.querySelector('em')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { fontWeight: '400', fontStyle: 'normal' }],
      [strong, { fontWeight: '700', fontStyle: 'normal' }],
      [em, { fontWeight: '400', fontStyle: 'italic' }],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any

    const boldRun = el.runs.find((r: any) => r.text === 'Bold')
    expect(boldRun?.bold).toBe(true)

    const italicRun = el.runs.find((r: any) => r.text === 'Italic')
    expect(italicRun?.italic).toBe(true)

    restore()
  })

  it('records hyperlink href', () => {
    const { section } = setupSlide(
      '<p id="t"><a href="https://example.com">Link</a></p>',
    )
    const p = document.getElementById('t')!
    const a = p.querySelector('a')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {}],
      [a, {}],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.runs).toHaveLength(1)
    expect(el.runs[0].text).toBe('Link')
    expect(el.runs[0].hyperlink).toContain('example.com')

    restore()
  })

  it('converts <br> to breakLine:true run', () => {
    const { section } = setupSlide('<p id="t">Line1<br>Line2</p>')
    const p = document.getElementById('t')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {}],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    // breakLine run should exist between Line1 and Line2
    const breakRun = el.runs.find((r: any) => r.breakLine === true)
    expect(breakRun).toBeDefined()
    expect(breakRun.text).toBe('')
    // Text runs should not contain literal '\n'
    const textRuns = el.runs.filter((r: any) => !r.breakLine)
    expect(textRuns.map((r: any) => r.text)).toEqual(['Line1', 'Line2'])

    restore()
  })
})

// -----------------------------------------------------------------------
// extractTextStyle — tested through extractSlides
// -----------------------------------------------------------------------

describe('extractTextStyle (via extractSlides)', () => {
  it('extracts CSS text style properties', () => {
    const { section } = setupSlide('<h1>Styled</h1>')
    const h1 = section.querySelector('h1')!

    mockRect(h1, { left: 0, top: 0, width: 600, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        h1,
        {
          color: 'rgb(34, 68, 102)',
          fontSize: '24px',
          fontFamily: '"Noto Sans JP", sans-serif',
          fontWeight: '700',
          textAlign: 'center',
          lineHeight: '36px',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.style).toEqual({
      color: 'rgb(34, 68, 102)',
      fontSize: 24,
      fontFamily: '"Noto Sans JP", sans-serif',
      fontWeight: 700,
      textAlign: 'center',
      lineHeight: 36,
      letterSpacing: 0,
    })

    restore()
  })
})

// -----------------------------------------------------------------------
// extractListItems — tested through extractSlides
// -----------------------------------------------------------------------

describe('extractListItems (via extractSlides)', () => {
  it('extracts flat list items', () => {
    const { section } = setupSlide(
      '<ul id="list"><li>Item A</li><li>Item B</li></ul>',
    )
    const ul = document.getElementById('list')!
    const lis = ul.querySelectorAll('li')

    mockRect(ul, { left: 0, top: 0, width: 600, height: 48 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ul, {}],
      [lis[0], {}],
      [lis[1], {}],
    ])

    const slides = extractSlides()
    const listEl = slides[0].elements[0] as any
    expect(listEl.items).toHaveLength(2)
    expect(listEl.items[0]).toMatchObject({ text: 'Item A', level: 0 })
    expect(listEl.items[1]).toMatchObject({ text: 'Item B', level: 0 })

    restore()
  })

  it('correctly records nested list levels', () => {
    const { section } = setupSlide(`
      <ul id="list">
        <li>Top
          <ul>
            <li>Nested</li>
          </ul>
        </li>
      </ul>
    `)
    const list = document.getElementById('list')!
    const allLis = list.querySelectorAll('li')

    mockRect(list, { left: 0, top: 0, width: 600, height: 72 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [list, {}],
      ...Array.from(allLis).map(
        (li) => [li, {}] as [Element, Record<string, string>],
      ),
    ])

    const slides = extractSlides()
    const listEl = slides[0].elements[0] as any
    expect(listEl.items.some((i: any) => i.level === 0)).toBe(true)
    expect(listEl.items.some((i: any) => i.level === 1)).toBe(true)

    restore()
  })

  it('preserves inline element order within list items', () => {
    const { section } = setupSlide(
      '<ul id="list"><li>Normal <strong>Bold</strong> more</li></ul>',
    )
    const list = document.getElementById('list')!
    const li = list.querySelector('li')!
    const strong = li.querySelector('strong')!

    mockRect(list, { left: 0, top: 0, width: 600, height: 48 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [list, {}],
      [li, { fontWeight: '400' }],
      [strong, { fontWeight: '700' }],
    ])

    const slides = extractSlides()
    const listEl = slides[0].elements[0] as any
    expect(listEl.items).toHaveLength(1)

    // Runs should be in document order: "Normal ", "Bold", " more"
    const texts = listEl.items[0].runs.map((r: any) => r.text)
    expect(texts[0]).toContain('Normal')
    expect(texts[1]).toBe('Bold')
    expect(texts[2]).toContain('more')

    // Bold run should have bold=true
    const boldRun = listEl.items[0].runs.find((r: any) => r.text === 'Bold')
    expect(boldRun.bold).toBe(true)

    restore()
  })

  it('correctly extracts li > p structured list items', () => {
    const { section } = setupSlide(
      '<ul id="list"><li><p>Paragraph in list</p></li></ul>',
    )
    const list = document.getElementById('list')!
    const li = list.querySelector('li')!
    const p = li.querySelector('p')!

    mockRect(list, { left: 0, top: 0, width: 600, height: 48 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [list, {}],
      [li, {}],
      [p, {}],
    ])

    const slides = extractSlides()
    const listEl = slides[0].elements[0] as any
    expect(listEl.items).toHaveLength(1)
    expect(listEl.items[0].runs.length).toBeGreaterThan(0)
    expect(listEl.items[0].runs[0].text).toBe('Paragraph in list')

    restore()
  })

  it('backgroundColor of inline strong element inside li propagates to text runs — slide 56/58 highlight', () => {
    const { section } = setupSlide(
      '<ul id="list"><li>Working on <strong id="s">development efficiency</strong> improvements</li></ul>',
    )
    const list = document.getElementById('list')!
    const li = list.querySelector('li')!
    const strong = document.getElementById('s')!

    mockRect(list, { left: 0, top: 0, width: 600, height: 48 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [list, {}],
      [li, {}],
      [strong, { display: 'inline', backgroundColor: 'rgb(241, 196, 15)' }],
    ])

    const slides = extractSlides()
    const listEl = slides[0].elements[0] as any
    const runs: any[] = listEl.items[0].runs

    // the 'development efficiency' run should have backgroundColor set
    const highlightRun = runs.find((r: any) => r.text === 'development efficiency')
    expect(highlightRun).toBeDefined()
    expect(highlightRun.backgroundColor).toBe('rgb(241, 196, 15)')

    // surrounding plain-text runs should have no backgroundColor
    const plainRuns = runs.filter((r: any) => r.text !== 'development efficiency' && !r.breakLine)
    plainRuns.forEach((r: any) => {
      expect(r.backgroundColor).toBeUndefined()
    })

    restore()
  })
})

// -----------------------------------------------------------------------
// extractTableData — tested through extractSlides
// -----------------------------------------------------------------------

describe('extractTableData (via extractSlides)', () => {
  it('extracts table rows and cells', () => {
    const { section } = setupSlide(`
      <table id="tbl">
        <tr><th>Header</th></tr>
        <tr><td>Cell</td></tr>
      </table>
    `)
    const table = document.getElementById('tbl')!
    const cells = table.querySelectorAll('th, td')

    mockRect(table, { left: 0, top: 0, width: 600, height: 80 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [table, {}],
      ...Array.from(cells).map(
        (c) => [c, {}] as [Element, Record<string, string>],
      ),
    ])

    const slides = extractSlides()
    const tableEl = slides[0].elements[0] as any
    expect(tableEl.rows).toHaveLength(2)
    expect(tableEl.rows[0].cells[0].isHeader).toBe(true)
    expect(tableEl.rows[0].cells[0].text).toBe('Header')
    expect(tableEl.rows[1].cells[0].isHeader).toBe(false)
    expect(tableEl.rows[1].cells[0].text).toBe('Cell')

    restore()
  })

  it('extracts inline decorations in table cells as runs', () => {
    const { section } = setupSlide(`
      <table id="tbl">
        <tr><td>Normal <strong>Bold</strong></td></tr>
      </table>
    `)
    const table = document.getElementById('tbl')!
    const td = table.querySelector('td')!
    const strong = td.querySelector('strong')!

    mockRect(table, { left: 0, top: 0, width: 600, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [table, {}],
      [td, { fontWeight: '400' }],
      [strong, { fontWeight: '700' }],
    ])

    const slides = extractSlides()
    const tableEl = slides[0].elements[0] as any
    const cell = tableEl.rows[0].cells[0]
    expect(cell.runs.length).toBeGreaterThan(0)

    const boldRun = cell.runs.find((r: any) => r.text === 'Bold')
    expect(boldRun?.bold).toBe(true)

    restore()
  })

  it('extracts table cell fontFamily', () => {
    const { section } = setupSlide(`
      <table id="tbl">
        <tr><td>Text</td></tr>
      </table>
    `)
    const table = document.getElementById('tbl')!
    const td = table.querySelector('td')!

    mockRect(table, { left: 0, top: 0, width: 600, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [table, {}],
      [td, { fontFamily: '"Noto Sans JP", sans-serif' }],
    ])

    const slides = extractSlides()
    const tableEl = slides[0].elements[0] as any
    expect(tableEl.rows[0].cells[0].style.fontFamily).toBe(
      '"Noto Sans JP", sans-serif',
    )

    restore()
  })

  it('extracts first-row cell widths as colWidths', () => {
    const { section } = setupSlide(`
      <table id="tbl">
        <tr><th id="c1">A</th><th id="c2">B</th></tr>
        <tr><td>1</td><td>2</td></tr>
      </table>
    `)
    const table = document.getElementById('tbl')!
    const c1 = document.getElementById('c1')!
    const c2 = document.getElementById('c2')!

    // mock offsetWidth
    Object.defineProperty(c1, 'offsetWidth', { value: 200, configurable: true })
    Object.defineProperty(c2, 'offsetWidth', { value: 400, configurable: true })

    mockRect(table, { left: 0, top: 0, width: 600, height: 60 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [table, {}],
      [c1, {}],
      [c2, {}],
      ...Array.from(table.querySelectorAll('td')).map(
        (c) => [c, {}] as [Element, Record<string, string>],
      ),
    ])

    const slides = extractSlides()
    const tableEl = slides[0].elements[0] as any
    expect(tableEl.colWidths).toEqual([200, 400])

    restore()
  })
})

// -----------------------------------------------------------------------
// Marp Inline SVG mode — section deduplication
// -----------------------------------------------------------------------

describe('Marp Inline SVG mode section deduplication', () => {
  it('merges 3-layer sections into 1 slide when ![bg] is used', () => {
    // Simulate Marp's 3-layer structure for a slide with ![bg]
    document.body.innerHTML = `
      <section data-marpit-pagination="1" data-marpit-advanced-background="background">
        <div data-marpit-advanced-background-container>
          <figure id="bg-figure"></figure>
        </div>
      </section>
      <section data-marpit-pagination="1" data-marpit-advanced-background="content">
        <h1>Title</h1>
      </section>
      <section data-marpit-pagination="1" data-marpit-advanced-background="pseudo">
      </section>
    `
    const sections = document.querySelectorAll('section')
    const h1 = document.querySelector('h1')!
    const figure = document.getElementById('bg-figure')!
    const bgContainer = document.querySelector(
      '[data-marpit-advanced-background-container]',
    )!

    for (const s of Array.from(sections)) {
      mockRect(s, { left: 0, top: 0, width: 1280, height: 720 })
    }
    mockRect(h1, { left: 10, top: 20, width: 500, height: 40 })

    const originalCS = globalThis.getComputedStyle
    const styleMap = new Map<Element, Record<string, string>>()
    for (const s of Array.from(sections)) {
      styleMap.set(s, {
        ...defaultStyles,
        backgroundColor: 'rgb(255,255,255)',
        backgroundImage: 'none',
      })
    }
    styleMap.set(h1, { ...defaultStyles, fontSize: '32px', fontWeight: '700' })
    styleMap.set(figure, {
      ...defaultStyles,
      backgroundImage: 'url("bg-image.png")',
      filter: 'none',
    })
    styleMap.set(bgContainer, { ...defaultStyles })
    ;(globalThis as any).getComputedStyle = (target: Element) => {
      const styles = styleMap.get(target) ?? defaultStyles
      return new Proxy({} as CSSStyleDeclaration, {
        get(_t, prop: string) {
          if (prop === 'getPropertyValue')
            return (name: string) => styles[name] ?? ''
          return styles[prop] ?? ''
        },
      })
    }

    const slides = extractSlides()

    // Should produce exactly 1 slide, not 3
    expect(slides).toHaveLength(1)
    expect(slides[0].sourceHasPagination).toBe(true)
    expect(slides[0].elements.some((e: any) => e.type === 'heading')).toBe(true)
    // Background image should be extracted from the figure as BgImageData[]
    expect(slides[0].backgroundImages).toHaveLength(1)
    expect(slides[0].backgroundImages[0].url).toBe('bg-image.png')
    ;(globalThis as any).getComputedStyle = originalCS
  })

  it('keeps 1 section = 1 slide without ![bg]', () => {
    const { section } = setupSlide('<p>Text</p>')
    const p = section.querySelector('p')!
    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {}],
    ])

    const slides = extractSlides()
    expect(slides).toHaveLength(1)

    restore()
  })
})

// -----------------------------------------------------------------------
// findBackgroundColor — tested through extractSlides
// -----------------------------------------------------------------------

describe('findBackgroundColor (via extractSlides)', () => {
  it('falls back to white when section is transparent', () => {
    const { section } = setupSlide('<p>Text</p>')
    const p = section.querySelector('p')!
    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })

    const restore = mockStyles([
      [
        section,
        { backgroundColor: 'rgba(0, 0, 0, 0)', backgroundImage: 'none' },
      ],
      [p, {}],
    ])

    const slides = extractSlides()
    expect(slides[0].background).toBe('rgb(255, 255, 255)')

    restore()
  })

  it('uses section background color as-is', () => {
    const { section } = setupSlide('<p>Text</p>')
    const p = section.querySelector('p')!
    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(30, 60, 90)' }],
      [p, {}],
    ])

    const slides = extractSlides()
    expect(slides[0].background).toBe('rgb(30, 60, 90)')

    restore()
  })

  it('extracts last opaque color from gradient background', () => {
    const { section } = setupSlide('<p>Text</p>')
    const p = section.querySelector('p')!
    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })

    const restore = mockStyles([
      [
        section,
        {
          backgroundColor: 'rgba(0, 0, 0, 0)',
          backgroundImage:
            'linear-gradient(135deg, rgba(15, 108, 189, 0.09), rgba(15, 108, 189, 0) 45%), linear-gradient(rgb(245, 251, 255) 0%, rgb(255, 255, 255) 75%)',
        },
      ],
      [p, {}],
    ])

    const slides = extractSlides()
    // Should extract the last rgb from the gradient: rgb(255, 255, 255)
    expect(slides[0].background).toBe('rgb(255, 255, 255)')

    restore()
  })

  it('gets correct color from section gradient even with black body', () => {
    const { section } = setupSlide('<p>Text</p>')
    const p = section.querySelector('p')!
    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })

    // Mock body with black background (just like Marp HTML)
    const originalBodyGetCS = globalThis.getComputedStyle
    const bodyProxy = new Proxy({} as CSSStyleDeclaration, {
      get(_t, prop: string) {
        if (prop === 'backgroundColor') return 'rgb(0, 0, 0)'
        return ''
      },
    })

    const sectionStyles = {
      ...defaultStyles,
      backgroundColor: 'rgba(0, 0, 0, 0)',
      backgroundImage:
        'linear-gradient(rgb(245, 251, 255) 0%, rgb(255, 255, 255) 75%)',
    }
    const sectionProxy = new Proxy({} as CSSStyleDeclaration, {
      get(_t, prop: string) {
        if (prop === 'getPropertyValue')
          return (name: string) => sectionStyles[name] ?? ''
        return sectionStyles[prop] ?? ''
      },
    })

    const pStyles = { ...defaultStyles }
    const pProxy = new Proxy({} as CSSStyleDeclaration, {
      get(_t, prop: string) {
        if (prop === 'getPropertyValue')
          return (name: string) => pStyles[name] ?? ''
        return pStyles[prop] ?? ''
      },
    })

    const map = new Map<Element, CSSStyleDeclaration>()
    map.set(section, sectionProxy)
    map.set(p, pProxy)
    ;(globalThis as any).getComputedStyle = (target: Element) => {
      if (target === document.body) return bodyProxy
      return map.get(target) ?? originalBodyGetCS(target)
    }

    const slides = extractSlides()
    // Should NOT pick up body's black background
    expect(slides[0].background).not.toBe('rgb(0, 0, 0)')
    expect(slides[0].background).toBe('rgb(255, 255, 255)')
    ;(globalThis as any).getComputedStyle = originalBodyGetCS
  })
})

// -----------------------------------------------------------------------
// blockquote border extraction
// -----------------------------------------------------------------------

describe('blockquote border (via extractSlides)', () => {
  it('extracts left border width and color', () => {
    const { section } = setupSlide('<blockquote>Quote text</blockquote>')
    const bq = section.querySelector('blockquote')!

    mockRect(bq, { left: 10, top: 50, width: 600, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [bq, { borderLeftWidth: '4px', borderLeftColor: 'rgb(100, 100, 100)' }],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.type).toBe('blockquote')
    expect(el.borderLeft).toEqual({ width: 4, color: 'rgb(100, 100, 100)' })

    restore()
  })

  it('omits borderLeft when no left border exists', () => {
    const { section } = setupSlide('<blockquote>Quote text</blockquote>')
    const bq = section.querySelector('blockquote')!

    mockRect(bq, { left: 10, top: 50, width: 600, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [bq, { borderLeftWidth: '0px' }],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.borderLeft).toBeUndefined()

    restore()
  })

  it('inserts breakLine run between multiple <p> elements in blockquote', () => {
    const { section } = setupSlide(
      '<blockquote id="bq"><p>First</p><p>Second</p></blockquote>',
    )
    const bq = document.getElementById('bq')!
    const paras = bq.querySelectorAll('p')

    mockRect(bq, { left: 10, top: 50, width: 600, height: 80 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [bq, { borderLeftWidth: '4px', borderLeftColor: 'rgb(50,50,50)' }],
      [paras[0], { display: 'block' }],
      [paras[1], { display: 'block' }],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    // Expect: [run'First', breakLine, run'Second'] (no trailing break)
    const texts = el.runs.map((r: any) => (r.breakLine ? '\n' : r.text))
    expect(texts).toEqual(['First', '\n', 'Second'])

    restore()
  })
})

// -----------------------------------------------------------------------
// code syntax highlighting
// -----------------------------------------------------------------------

describe('code syntax highlighting (via extractSlides)', () => {
  it('extracts colored runs from span elements inside pre > code', () => {
    const { section } = setupSlide(
      '<pre id="codeblock"><code><span class="keyword">const</span> x = 1;</code></pre>',
    )
    const pre = document.getElementById('codeblock')!
    const code = pre.querySelector('code')!
    const span = code.querySelector('span')!

    mockRect(pre, { left: 10, top: 100, width: 600, height: 80 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [pre, { backgroundColor: 'rgb(40, 44, 52)', fontSize: '14px' }],
      [code, { color: 'rgb(200, 200, 200)', fontSize: '14px' }],
      [
        span,
        { color: 'rgb(198, 120, 221)', fontSize: '14px', fontWeight: '700' },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.type).toBe('code')
    expect(el.runs.length).toBeGreaterThan(0)

    const keywordRun = el.runs.find((r: any) => r.text === 'const')
    expect(keywordRun).toBeDefined()
    expect(keywordRun.color).toBe('rgb(198, 120, 221)')
    expect(keywordRun.bold).toBe(true)

    restore()
  })
})

// -----------------------------------------------------------------------
// extractTextRuns — background-color (mark / strong / span)
// -----------------------------------------------------------------------

describe('extractTextRuns — backgroundColor (via extractSlides)', () => {
  it('propagates <mark> background-color to TextRun', () => {
    const { section } = setupSlide(
      '<p id="t">Normal <mark>Highlighted</mark> text</p>',
    )
    const p = document.getElementById('t')!
    const mark = p.querySelector('mark')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { display: 'block' }],
      [mark, { display: 'inline', backgroundColor: 'rgb(241, 196, 15)' }],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    const hlRun = el.runs.find((r: any) => r.text === 'Highlighted')
    expect(hlRun).toBeDefined()
    expect(hlRun.backgroundColor).toBe('rgb(241, 196, 15)')

    restore()
  })

  it('inline elements without background-color have no backgroundColor', () => {
    const { section } = setupSlide('<p id="t">Normal <em>Italic</em> text</p>')
    const p = document.getElementById('t')!
    const em = p.querySelector('em')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { display: 'block' }],
      [
        em,
        {
          display: 'inline',
          fontStyle: 'italic',
          backgroundColor: 'rgba(0, 0, 0, 0)',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    const italicRun = el.runs.find((r: any) => r.text === 'Italic')
    expect(italicRun).toBeDefined()
    expect(italicRun.backgroundColor).toBeUndefined()

    restore()
  })

  it('preserves independent backgroundColor for multiple <mark> elements', () => {
    const { section } = setupSlide(
      '<p id="t"><mark>A</mark> and <mark>B</mark></p>',
    )
    const p = document.getElementById('t')!
    const marks = p.querySelectorAll('mark')

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { display: 'block' }],
      [marks[0], { display: 'inline', backgroundColor: 'rgb(241, 196, 15)' }],
      [marks[1], { display: 'inline', backgroundColor: 'rgb(52, 152, 219)' }],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    const runA = el.runs.find((r: any) => r.text === 'A')
    const runB = el.runs.find((r: any) => r.text === 'B')
    expect(runA?.backgroundColor).toBe('rgb(241, 196, 15)')
    expect(runB?.backgroundColor).toBe('rgb(52, 152, 219)')

    restore()
  })
})

// -----------------------------------------------------------------------
// extractTextRuns — linear-gradient backgroundImage as approximate highlight
// Regression for slide 42: .marker-highlight strong uses
// background: linear-gradient(transparent 62%, #fff2a8 62%)
// The last solid colour stop should be extracted as backgroundColor.
// -----------------------------------------------------------------------

describe('extractTextRuns — linear-gradient backgroundImage (via extractSlides)', () => {
  it('extracts last solid color from two-stop transparent→color gradient as backgroundColor', () => {
    const { section } = setupSlide('<p id="t">Normal <strong id="hl">highlight</strong> text</p>')
    const p = document.getElementById('t')!
    const strong = document.getElementById('hl')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { display: 'block' }],
      [
        strong,
        {
          display: 'inline',
          // Chromium computes the gradient as resolved rgb() values
          backgroundImage:
            'linear-gradient(rgba(0, 0, 0, 0) 62%, rgb(255, 242, 168) 62%)',
          backgroundColor: 'rgba(0, 0, 0, 0)',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    const hlRun = el.runs.find((r: any) => r.text === 'highlight')
    expect(hlRun).toBeDefined()
    expect(hlRun.backgroundColor).toBe('rgb(255, 242, 168)')

    restore()
  })

  it('skips gradient when all stops are transparent', () => {
    const { section } = setupSlide('<p id="t">Text <em id="em">em</em> more</p>')
    const p = document.getElementById('t')!
    const em = document.getElementById('em')!

    mockRect(p, { left: 0, top: 0, width: 600, height: 24 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { display: 'block' }],
      [
        em,
        {
          display: 'inline',
          backgroundImage:
            'linear-gradient(rgba(0, 0, 0, 0) 62%, rgba(0, 0, 0, 0) 62%)',
          backgroundColor: 'rgba(0, 0, 0, 0)',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    const emRun = el.runs.find((r: any) => r.text === 'em')
    expect(emRun?.backgroundColor).toBeUndefined()

    restore()
  })
})

// -----------------------------------------------------------------------
// heading border extraction
// -----------------------------------------------------------------------

describe('heading border extraction (via extractSlides)', () => {
  it('extracts h1 border-bottom as borderBottom', () => {
    const { section } = setupSlide('<h1>Title</h1>')
    const h1 = section.querySelector('h1')!

    mockRect(h1, { left: 0, top: 0, width: 600, height: 50 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        h1,
        {
          fontSize: '40px',
          fontWeight: '700',
          borderBottomWidth: '2px',
          borderBottomColor: 'rgb(39, 174, 96)',
          borderLeftWidth: '0px',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.type).toBe('heading')
    expect(el.borderBottom).toEqual({ width: 2, color: 'rgb(39, 174, 96)' })
    expect(el.borderLeft).toBeUndefined()

    restore()
  })

  it('extracts h2 border-left as borderLeft', () => {
    const { section } = setupSlide('<h2>Section</h2>')
    const h2 = section.querySelector('h2')!

    mockRect(h2, { left: 0, top: 0, width: 600, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        h2,
        {
          fontSize: '30px',
          fontWeight: '700',
          borderLeftWidth: '4px',
          borderLeftColor: 'rgb(39, 174, 96)',
          borderBottomWidth: '0px',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.type).toBe('heading')
    expect(el.borderLeft).toEqual({ width: 4, color: 'rgb(39, 174, 96)' })
    expect(el.borderBottom).toBeUndefined()

    restore()
  })

  it('omits border properties when both border-bottom and border-left are 0', () => {
    const { section } = setupSlide('<h1>Plain</h1>')
    const h1 = section.querySelector('h1')!

    mockRect(h1, { left: 0, top: 0, width: 600, height: 50 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        h1,
        {
          fontSize: '40px',
          fontWeight: '700',
          borderBottomWidth: '0px',
          borderLeftWidth: '0px',
        },
      ],
    ])

    const slides = extractSlides()
    const el = slides[0].elements[0] as any
    expect(el.borderBottom).toBeUndefined()
    expect(el.borderLeft).toBeUndefined()

    restore()
  })
})

// -----------------------------------------------------------------------
// SVG embedding
// -----------------------------------------------------------------------

describe('SVG element extraction (via extractSlides)', () => {
  it('extracts <svg> inside slide as image element', () => {
    const { section } = setupSlide(
      '<svg width="200" height="100" viewBox="0 0 200 100"><rect x="10" y="10" width="80" height="60" fill="blue"/></svg>',
    )
    const svg = section.querySelector('svg')!

    mockRect(svg, { left: 10, top: 20, width: 200, height: 100 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [svg, { display: 'inline', visibility: 'visible' }],
    ])

    ;(svg as any).getBoundingClientRect = () => ({
      left: 10,
      top: 20,
      width: 200,
      height: 100,
      right: 210,
      bottom: 120,
    })

    const slides = extractSlides()
    const svgEl = slides[0].elements.find((e: any) => e.type === 'image') as any
    expect(svgEl).toBeDefined()
    // SVG is base64-encoded for PptxGenJS/Office compatibility
    expect(svgEl.src).toMatch(/^data:image\/svg\+xml;base64,/)

    restore()
  })
})

// -----------------------------------------------------------------------
// Inline children inside flex parent
// -----------------------------------------------------------------------

describe('preserving display:inline children in flex container (via extractSlides)', () => {
  it('extracts display:inline text span inside display:flex parent as paragraph', () => {
    const { section } = setupSlide(`
      <div id="flex-row">
        <span id="badge" style="border-radius:50%;">1</span>
        <span id="label">Text Label</span>
      </div>
    `)
    const flexDiv = section.querySelector('#flex-row')!
    const badge = section.querySelector('#badge')!
    const label = section.querySelector('#label')!

    mockRect(flexDiv, { left: 0, top: 10, width: 600, height: 40 })
    mockRect(badge, { left: 0, top: 14, width: 32, height: 32 })
    mockRect(label, { left: 42, top: 18, width: 200, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [flexDiv, { display: 'flex', alignItems: 'center' }],
      [
        badge,
        {
          display: 'inline',
          backgroundColor: 'rgb(0, 102, 204)',
          color: 'rgb(255,255,255)',
          borderRadius: '50%',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          textAlign: 'left',
          lineHeight: '14px',
        },
      ],
      [
        label,
        {
          display: 'inline',
          color: 'rgb(0,0,0)',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          textAlign: 'left',
          lineHeight: '24px',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
    ])

    const slides = extractSlides()
    const container = slides[0].elements.find(
      (e: any) => e.type === 'container',
    ) as any
    expect(container).toBeDefined()

    // Verify label span is also included as a paragraph in children
    const allElements = container?.children ?? slides[0].elements
    const paragraphs = allElements.filter((e: any) => e.type === 'paragraph')
    const labelParagraph = paragraphs.find((e: any) =>
      e.runs?.some((r: any) => r.text === 'Text Label'),
    )
    expect(labelParagraph).toBeDefined()

    restore()
  })

  it('display:inline-flex badge span has paragraph with valign:middle', () => {
    const { section } = setupSlide('<span id="badge">1</span>')
    const badge = section.querySelector('#badge')!

    mockRect(badge, { left: 10, top: 10, width: 32, height: 32 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        badge,
        {
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: 'rgb(0, 102, 204)',
          color: 'rgb(255,255,255)',
          borderRadius: '16px',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          textAlign: 'left',
          lineHeight: '14px',
        },
      ],
    ])

    const slides = extractSlides()
    const para = slides[0].elements.find(
      (e: any) => e.type === 'paragraph',
    ) as any
    expect(para).toBeDefined()
    expect(para.valign).toBe('middle')

    restore()
  })
})

// -----------------------------------------------------------------------
// extractListItems — <br> (trailing-space hard line break) inside tight list li
// -----------------------------------------------------------------------

describe('extractListItems — <br> inside tight list <li> (via extractSlides)', () => {
  it('produces a breakLine run for <br> inside a tight list item', () => {
    // Markdown:  - First line  \n  Second line
    // Rendered HTML: <li>First line<br>Second line</li>
    const { section } = setupSlide(
      '<ul id="ul"><li id="li">First line<br>Second line</li></ul>',
    )
    const ul = section.querySelector('#ul')!
    const li = section.querySelector('#li')!

    mockRect(ul, { left: 0, top: 0, width: 600, height: 48 })
    mockRect(li, { left: 0, top: 0, width: 600, height: 48 })
    const restore = mockStyles([
      [
        section,
        { backgroundColor: 'rgb(255,255,255)' },
      ],
      [
        ul,
        {
          display: 'block',
          fontSize: '16px',
          fontFamily: 'Arial',
          color: 'rgb(0,0,0)',
          fontWeight: '400',
          fontStyle: 'normal',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
      [
        li,
        {
          display: 'list-item',
          fontSize: '16px',
          fontFamily: 'Arial',
          color: 'rgb(0,0,0)',
          fontWeight: '400',
          fontStyle: 'normal',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
    ])

    const slides = extractSlides()
    const list = slides[0].elements.find((e: any) => e.type === 'list') as any
    expect(list).toBeDefined()

    const item = list.items[0]
    // Should have: "First line", breakLine, "Second line"
    expect(item.runs).toHaveLength(3)
    expect(item.runs[0]).toMatchObject({ text: 'First line' })
    expect(item.runs[1]).toMatchObject({ breakLine: true })
    expect(item.runs[2]).toMatchObject({ text: 'Second line' })

    restore()
  })
})

// -----------------------------------------------------------------------
// extractListItems — emoji img directly inside li (tight list, no <p> wrapper)
// -----------------------------------------------------------------------

describe('extractListItems — emoji img directly inside tight list li (via extractSlides)', () => {
  it('emoji img directly inside li is extracted as text run', () => {
    const { section } = setupSlide(
      '<ul id="ul"><li id="li">request<img class="emoji" alt="👉" src="https://twemoji/1f449.svg">analysis</li></ul>',
    )
    const ul = section.querySelector('#ul')!
    const li = section.querySelector('#li')!
    const img = section.querySelector('img')!

    mockRect(ul, { left: 0, top: 0, width: 600, height: 30 })
    mockRect(li, { left: 0, top: 0, width: 600, height: 30 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        ul,
        {
          display: 'block',
          fontSize: '16px',
          fontFamily: 'Arial',
          color: 'rgb(0,0,0)',
          fontWeight: '400',
          fontStyle: 'normal',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
      [
        li,
        {
          display: 'list-item',
          fontSize: '16px',
          fontFamily: 'Arial',
          color: 'rgb(0,0,0)',
          fontWeight: '400',
          fontStyle: 'normal',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
      [
        img,
        {
          display: 'inline',
          fontSize: '16px',
          fontFamily: 'Arial',
          color: 'rgb(0,0,0)',
          fontWeight: '400',
          fontStyle: 'normal',
          textAlign: 'left',
          lineHeight: '24px',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
    ])

    const slides = extractSlides()
    const list = slides[0].elements.find((e: any) => e.type === 'list') as any
    expect(list).toBeDefined()

    const item = list.items[0]
    const texts = item.runs
      .filter((r: any) => !r.breakLine)
      .map((r: any) => r.text)
    expect(texts).toContain('request')
    expect(texts).toContain('👉')
    expect(texts).toContain('analysis')

    restore()
  })

  it('extracts consecutive emoji as multiple text runs', () => {
    const { section } = setupSlide(
      '<ul id="ul"><li id="li">A<img class="emoji" alt="👉" src="https://twemoji/1f449.svg">B<img class="emoji" alt="👉" src="https://twemoji/1f449.svg">C</li></ul>',
    )
    const ul = section.querySelector('#ul')!
    const li = section.querySelector('#li')!
    const imgs = Array.from(section.querySelectorAll('img'))

    mockRect(ul, { left: 0, top: 0, width: 600, height: 30 })
    mockRect(li, { left: 0, top: 0, width: 600, height: 30 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        ul,
        {
          display: 'block',
          fontSize: '16px',
          fontFamily: 'Arial',
          color: 'rgb(0,0,0)',
          fontWeight: '400',
          fontStyle: 'normal',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
      [
        li,
        {
          display: 'list-item',
          fontSize: '16px',
          fontFamily: 'Arial',
          color: 'rgb(0,0,0)',
          fontWeight: '400',
          fontStyle: 'normal',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
      ...imgs.map(
        (img) =>
          [
            img,
            {
              display: 'inline',
              fontSize: '16px',
              fontFamily: 'Arial',
              color: 'rgb(0,0,0)',
              fontWeight: '400',
              fontStyle: 'normal',
              textAlign: 'left',
              lineHeight: '24px',
              backgroundColor: 'rgba(0,0,0,0)',
            },
          ] as [Element, Record<string, string>],
      ),
    ])

    const slides = extractSlides()
    const list = slides[0].elements.find((e: any) => e.type === 'list') as any
    const item = list.items[0]
    const texts = item.runs
      .filter((r: any) => !r.breakLine)
      .map((r: any) => r.text)
    expect(texts.filter((t: string) => t === '👉')).toHaveLength(2)
    expect(texts).toContain('A')
    expect(texts).toContain('B')
    expect(texts).toContain('C')

    restore()
  })
})

// -----------------------------------------------------------------------
// extractTextStyle — justify-content:center → textAlign:center
// -----------------------------------------------------------------------

describe('extractTextStyle justify-content:center mapping (via extractSlides)', () => {
  it('justify-content:center span has paragraph with textAlign:center', () => {
    const { section } = setupSlide('<span id="badge">1</span>')
    const badge = section.querySelector('#badge')!

    mockRect(badge, { left: 10, top: 10, width: 32, height: 32 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        badge,
        {
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          textAlign: 'left', // CSS textAlign is left, justifyContent drives centering
          backgroundColor: 'rgb(0, 102, 204)',
          color: 'rgb(255,255,255)',
          borderRadius: '16px',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          lineHeight: '14px',
        },
      ],
    ])

    const slides = extractSlides()
    const para = slides[0].elements.find(
      (e: any) => e.type === 'paragraph',
    ) as any
    expect(para).toBeDefined()
    expect(para.style.textAlign).toBe('center')

    restore()
  })
})

// -----------------------------------------------------------------------
// extractInlineBadgeShapes — inline-block badge / pill inside paragraph
// -----------------------------------------------------------------------

describe('extractInlineBadgeShapes — inline-block badge inside paragraph (via extractSlides)', () => {
  it('leading badge in <p>: container shape emitted, paragraph starts after badge, no highlight', () => {
    // <p><badge>01</badge> Step title</p> — badge is at para left edge (leading)
    // → computeLeadingOffset → shape emitted, para x shifted right by badge width
    const { section } = setupSlide(`
      <p id="para"><span id="badge">01</span> Step title</p>
    `)
    const para = section.querySelector('#para')!
    const badge = section.querySelector('#badge')!

    mockRect(para, { left: 50, top: 100, width: 600, height: 36 })
    mockRect(badge, { left: 50, top: 102, width: 50, height: 32 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        para,
        {
          display: 'block',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          color: 'rgb(0,0,0)',
          lineHeight: '22px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [
        badge,
        {
          display: 'inline-block',
          backgroundColor: 'rgb(0,102,204)',
          color: 'rgb(255,255,255)',
          borderRadius: '999px',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          textAlign: 'center',
          lineHeight: '14px',
        },
      ],
    ])

    const slides = extractSlides()
    const elements = slides[0].elements

    // Container shape MUST be emitted (leading badge → always a shape)
    const containerIdx = elements.findIndex((e: any) => e.type === 'container')
    const paragraphIdx = elements.findIndex((e: any) => e.type === 'paragraph')
    expect(containerIdx).toBeGreaterThanOrEqual(0)
    expect(paragraphIdx).toBeGreaterThan(containerIdx)

    // Badge text is in the shape — with NO backgroundColor (no stray highlight)
    const container = elements[containerIdx] as any
    expect(container.style.backgroundColor).toBe('rgb(0,102,204)')
    expect(container.style.borderRadius).toBe(999)
    const badgeRunInShape = container.runs?.find((r: any) => r.text === '01')
    expect(badgeRunInShape).toBeDefined()
    expect(badgeRunInShape.backgroundColor).toBeUndefined() // KEY: no highlight on shape text

    // Paragraph text box is shifted right to start after the badge
    // para.left=50, badge.width=50 → para x = 50+50=100, width = 600-50=550
    const paragraph = elements[paragraphIdx] as any
    expect(paragraph.x).toBe(100) // offset by badge width
    expect(paragraph.width).toBe(550) // reduced by badge width

    // Badge text '01' NOT in paragraph runs
    const badgeInPara = paragraph.runs?.find((r: any) => r.text === '01')
    expect(badgeInPara).toBeUndefined()

    restore()
  })

  it('ISOLATED badge in <p>: emits container shape with runs, no paragraph', () => {
    // <p><badge>HIGH</badge></p> — badge is the sole content (no text nodes)
    // → container shape with centered text, no paragraph emitted
    const { section } = setupSlide(`
      <p id="para"><span id="badge">HIGH</span></p>
    `)
    const para = section.querySelector('#para')!
    const badge = section.querySelector('#badge')!

    mockRect(para, { left: 50, top: 100, width: 120, height: 36 })
    mockRect(badge, { left: 50, top: 102, width: 100, height: 32 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        para,
        {
          display: 'block',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          color: 'rgb(0,0,0)',
          lineHeight: '22px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [
        badge,
        {
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: 'rgb(220,53,69)',
          color: 'rgb(255,255,255)',
          borderRadius: '999px',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          textAlign: 'center',
          lineHeight: '14px',
        },
      ],
    ])

    const slides = extractSlides()
    const elements = slides[0].elements

    // Isolated badge (no surrounding text) → shape only, no paragraph
    const containerIdx = elements.findIndex((e: any) => e.type === 'container')
    expect(containerIdx).toBeGreaterThanOrEqual(0)
    const container = elements[containerIdx] as any
    expect(container.style.backgroundColor).toBe('rgb(220,53,69)')
    const badgeRun = container.runs?.find((r: any) => r.text === 'HIGH')
    expect(badgeRun).toBeDefined()
    expect(badgeRun.backgroundColor).toBeUndefined() // no highlight on shape text

    // No paragraph (no non-badge text → extractTextRuns returns empty)
    const paragraphIdx = elements.findIndex((e: any) => e.type === 'paragraph')
    expect(paragraphIdx).toBe(-1)

    restore()
  })

  it('leading badge in heading: container shape emitted, heading starts after badge, no highlight', () => {
    // <h2><badge>01</badge> Section Title</h2> — badge at heading left edge
    // → shape emitted, heading text box shifted right by badge width
    const { section } = setupSlide(`
      <h2 id="heading"><span id="badge">01</span> Section Title</h2>
    `)
    const heading = section.querySelector('#heading')!
    const badge = section.querySelector('#badge')!

    mockRect(heading, { left: 50, top: 50, width: 700, height: 48 })
    mockRect(badge, { left: 50, top: 55, width: 44, height: 38 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        heading,
        {
          display: 'block',
          fontSize: '32px',
          fontFamily: 'Arial',
          fontWeight: '700',
          color: 'rgb(0,0,0)',
          lineHeight: '40px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
          borderBottomWidth: '0px',
          borderLeftWidth: '0px',
        },
      ],
      [
        badge,
        {
          display: 'inline-block',
          backgroundColor: 'rgb(200,50,50)',
          color: 'rgb(255,255,255)',
          borderRadius: '999px',
          fontSize: '18px',
          fontFamily: 'Arial',
          fontWeight: '700',
          textAlign: 'center',
          lineHeight: '18px',
        },
      ],
    ])

    const slides = extractSlides()
    const elements = slides[0].elements

    // Container for badge emitted before the heading
    const containerIdx = elements.findIndex((e: any) => e.type === 'container')
    const headingIdx = elements.findIndex((e: any) => e.type === 'heading')
    expect(containerIdx).toBeGreaterThanOrEqual(0)
    expect(headingIdx).toBeGreaterThan(containerIdx)

    const container = elements[containerIdx] as any
    expect(container.style.backgroundColor).toBe('rgb(200,50,50)')
    expect(container.style.borderRadius).toBe(999)
    const badgeRunInShape = container.runs?.find((r: any) => r.text === '01')
    expect(badgeRunInShape).toBeDefined()
    expect(badgeRunInShape.backgroundColor).toBeUndefined() // no highlight on shape text

    // Heading starts after the badge: x = 50+44=94, width = 700-44=656
    const headingEl = elements[headingIdx] as any
    expect(headingEl.x).toBe(94)
    expect(headingEl.width).toBe(656)

    // Badge text '01' NOT in heading runs
    const badgeInHeading = headingEl.runs?.find((r: any) => r.text === '01')
    expect(badgeInHeading).toBeUndefined()

    restore()
  })

  it('does not emit container for inline span without background', () => {
    const { section } = setupSlide(
      '<p id="para"><span id="normal">plain text</span></p>',
    )
    const para = section.querySelector('#para')!
    const span = section.querySelector('#normal')!

    mockRect(para, { left: 0, top: 0, width: 600, height: 24 })
    mockRect(span, { left: 0, top: 0, width: 100, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        para,
        {
          display: 'block',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          color: 'rgb(0,0,0)',
          lineHeight: '24px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [
        span,
        {
          display: 'inline-block',
          backgroundColor: 'rgba(0,0,0,0)',
          color: 'rgb(0,0,0)',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
    ])

    const slides = extractSlides()
    const containerEl = slides[0].elements.find(
      (e: any) => e.type === 'container',
    )
    expect(containerEl).toBeUndefined()

    restore()
  })

  it('非 leading の inline-flex バッジは bg-only container シェイプになり、テキストは backgroundColor なしで段落に残る', () => {
    // badge.left=200, para.left=50 → 200 > 50+8 → non-leading
    // NEW behavior: background-only container shape emitted;
    //   badge text stays in paragraph run WITHOUT backgroundColor.
    const { section } = setupSlide(`
      <p id="para">Step <span id="badge">MID</span> flow</p>
    `)
    const para = section.querySelector('#para')!
    const badge = section.querySelector('#badge')!

    mockRect(para, { left: 50, top: 100, width: 600, height: 36 })
    mockRect(badge, { left: 200, top: 104, width: 40, height: 28 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        para,
        {
          display: 'block',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          color: 'rgb(0,0,0)',
          lineHeight: '22px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [
        badge,
        {
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: 'rgb(0,102,204)',
          color: 'rgb(255,255,255)',
          borderRadius: '999px',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          textAlign: 'center',
          lineHeight: '14px',
        },
      ],
    ])

    const slides = extractSlides()
    const elements = slides[0].elements

    // 非 leading → bg-only container シェイプが出力される (runs なし)
    const containerEl = elements.find((e: any) => e.type === 'container') as any
    expect(containerEl).toBeDefined()
    expect(containerEl.style.backgroundColor).toBe('rgb(0,102,204)')
    // bg-only shape has no text runs
    const bgOnlyHasText = containerEl.runs?.some((r: any) => !r.breakLine && r.text?.trim() !== '')
    expect(bgOnlyHasText).toBeFalsy()

    // 段落が出力され、バッジテキストは文字色を維持するが backgroundColor は付かない
    const paraEl = elements.find((e: any) => e.type === 'paragraph') as any
    expect(paraEl).toBeDefined()
    const badgeRun = paraEl.runs?.find((r: any) => r.text === 'MID')
    expect(badgeRun).toBeDefined()
    // backgroundColor が剥ぎ取られている (bg-only shape が視覚的背景を提供する)
    expect(badgeRun.backgroundColor).toBeUndefined()

    restore()
  })
})

// -----------------------------------------------------------------------
// Highlight bleed fix — inline-only div with background-color
// Regression for slides 30/48: when a <div> has a background-color and
// only inline/text content, dom-walker emits a container shape + paragraph.
// The paragraph runs must NOT carry backgroundColor matching the container
// (it would cause colour bleed when text drifts slightly from the shape).
// -----------------------------------------------------------------------

describe('inline-only div: run backgroundColor stripped when container provides background', () => {
  function setupInlineDiv(html: string) {
    const { section } = setupSlide(html)
    return section
  }

  it('strips element background-color from runs inside an inline-only coloured div', () => {
    const section = setupInlineDiv('<div id="box">Blue label</div>')
    const box = section.querySelector('#box')!

    mockRect(box, { left: 10, top: 10, width: 200, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        box,
        {
          display: 'block',
          backgroundColor: 'rgb(26, 115, 232)',
          color: 'rgb(255,255,255)',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          fontStyle: 'normal',
          textDecorationLine: 'none',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
    ])

    const slides = extractSlides()
    const elements = slides[0].elements
    // Container shape should exist
    const containerEl = elements.find((e: any) => e.type === 'container') as any
    expect(containerEl).toBeDefined()
    expect(containerEl.style.backgroundColor).toBe('rgb(26, 115, 232)')
    // Paragraph runs must not have backgroundColor equal to the container fill
    const paraEl = elements.find((e: any) => e.type === 'paragraph') as any
    expect(paraEl).toBeDefined()
    const run = paraEl.runs.find((r: any) => r.text === 'Blue label')
    expect(run).toBeDefined()
    expect(run.backgroundColor).toBeUndefined()

    restore()
  })

  it('preserves genuine inline highlight (different color) inside a coloured div', () => {
    const section = setupInlineDiv(
      '<div id="box">Normal <mark id="hl">important</mark> text</div>',
    )
    const box = section.querySelector('#box')!
    const mark = section.querySelector('#hl')!

    mockRect(box, { left: 10, top: 10, width: 300, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        box,
        {
          display: 'block',
          backgroundColor: 'rgb(26, 115, 232)',
          color: 'rgb(255,255,255)',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          fontStyle: 'normal',
          textDecorationLine: 'none',
          textAlign: 'left',
          lineHeight: '24px',
        },
      ],
      [
        mark,
        {
          display: 'inline',
          backgroundColor: 'rgb(255, 193, 7)',
          color: 'rgb(0, 0, 0)',
        },
      ],
    ])

    const slides = extractSlides()
    const elements = slides[0].elements
    const paraEl = elements.find((e: any) => e.type === 'paragraph') as any
    expect(paraEl).toBeDefined()
    // Direct text "Normal" / "text": no background (would have been container bg → stripped)
    const normalRun = paraEl.runs.find((r: any) => r.text === 'Normal')
    expect(normalRun?.backgroundColor).toBeUndefined()
    // Inline highlight with a DIFFERENT color must be preserved
    const hlRun = paraEl.runs.find((r: any) => r.text === 'important')
    expect(hlRun?.backgroundColor).toBe('rgb(255, 193, 7)')

    restore()
  })
})

// -----------------------------------------------------------------------
// Slide 48 regression: inline-block child inside a block container div
// must NOT produce a duplicate text element.
//
// Pattern:
//   <div style="text-align:center">         ← block container
//     <div style="display:inline-block">    ← badge with background
//       Input data
//     </div>
//   </div>
//
// walkElements processes the inner inline-block div (display is not 'inline'
// so it is not skipped), producing a container shape + paragraph.  The shallow
// walk must NOT then also extract the inner div's text into a second paragraph.
// -----------------------------------------------------------------------

describe('inline-block badge inside block container div — no duplicate text (slide 48)', () => {
  it('inner inline-block div produces exactly one paragraph (no duplicate)', () => {
    const { section } = setupSlide(
      '<div id="outer"><div id="badge">Input data</div></div>',
    )
    const outer = section.querySelector('#outer')!
    const badge = section.querySelector('#badge')!

    mockRect(outer, { left: 340, top: 280, width: 300, height: 44 })
    mockRect(badge, { left: 340, top: 280, width: 300, height: 44 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        outer,
        {
          display: 'block',
          textAlign: 'center',
          backgroundColor: 'rgba(0,0,0,0)',
          color: 'rgb(0,0,0)',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          fontStyle: 'normal',
          lineHeight: '24px',
        },
      ],
      [
        badge,
        {
          display: 'inline-block',
          backgroundColor: 'rgb(26, 115, 232)',
          color: 'rgb(255,255,255)',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '700',
          fontStyle: 'normal',
          lineHeight: '24px',
          textAlign: 'center',
          borderRadius: '8px',
        },
      ],
    ])

    const slides = extractSlides()

    // Flatten all elements (including nested container children)
    function flatten(els: any[]): any[] {
      return els.flatMap((e: any) =>
        e.children ? [e, ...flatten(e.children)] : [e],
      )
    }
    const all = flatten(slides[0].elements)

    // There should be exactly ONE paragraph containing "Input data"
    const paragraphs = all.filter(
      (e: any) =>
        e.type === 'paragraph' &&
        e.runs?.some((r: any) => r.text === 'Input data'),
    )
    expect(paragraphs).toHaveLength(1)

    restore()
  })
})

// -----------------------------------------------------------------------
// extractPseudoElements — content:'' decorative bars
//
// Rules:
//  - content:'' + section HAS user class → extract (e.g. section.decorated)
//    UNLESS the same background also appears on classless sections (global rule)
//  - content:'' + section has NO user class → skip (Marp scoped-style artifact)
//  - transparent background → skip regardless
//  - content:'' global rule (same bg on classless sections) → skip for all
// -----------------------------------------------------------------------

describe('extractPseudoElements (via extractSlides): content empty-string pseudo-element with background', () => {
  it('content:\'\' + section has user class → extracted as container bar', () => {
    const { section } = setupSlide('<p>Content</p>')
    const p = section.querySelector('p')!

    mockRect(p, { left: 0, top: 0, width: 640, height: 24 })
    // Give section a user class (e.g. "decorated")
    section.className = 'decorated'

    // Install element styles first, then wrap getComputedStyle to add pseudo-element handling
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { fontSize: '16px', fontWeight: '400', color: 'rgb(0,0,0)' }],
    ])

    const csWithStyles = (globalThis as any).getComputedStyle
    ;(globalThis as any).getComputedStyle = (el: Element, pseudo?: string | null) => {
      if (pseudo === '::before' && el === section) {
        return {
          content: '""',
          backgroundColor: 'rgb(37, 99, 235)',
          position: 'absolute',
          top: '0px',
          left: '0px',
          width: '1280px',
          height: '12px',
          display: 'block',
        } as any
      }
      if (pseudo === '::after' && el === section) {
        return { content: 'none', backgroundColor: 'rgba(0,0,0,0)' } as any
      }
      return csWithStyles(el, pseudo)
    }

    const slides = extractSlides()
    const bar = slides[0].elements.find(
      (e: any) => e.type === 'container' && e.y === 0 && e.height === 12,
    ) as any
    expect(bar).toBeDefined()
    expect(bar.style.backgroundColor).toBe('rgb(37, 99, 235)')

    ;(globalThis as any).getComputedStyle = csWithStyles
    section.className = ''
    restore()
  })

  it('content:\'\' + section has NO user class → skipped (Marp scoped-style artifact)', () => {
    const { section } = setupSlide('<p>Content</p>')
    const p = section.querySelector('p')!

    mockRect(p, { left: 0, top: 0, width: 640, height: 24 })
    // No user class

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { fontSize: '16px', fontWeight: '400', color: 'rgb(0,0,0)' }],
    ])

    const csWithStyles = (globalThis as any).getComputedStyle
    ;(globalThis as any).getComputedStyle = (el: Element, pseudo?: string | null) => {
      if (pseudo === '::before' && el === section) {
        return {
          content: '""',
          backgroundColor: 'rgb(15, 108, 189)',  // would appear as banner
          position: 'absolute',
          top: '0px',
          left: '0px',
          width: '1280px',
          height: '16px',
          display: 'block',
        } as any
      }
      if (pseudo === '::after' && el === section) {
        return { content: 'none', backgroundColor: 'rgba(0,0,0,0)' } as any
      }
      return csWithStyles(el, pseudo)
    }

    const slides = extractSlides()
    const bar = slides[0].elements.find(
      (e: any) => e.type === 'container' && e.y === 0 && e.height === 16,
    )
    expect(bar).toBeUndefined()

    ;(globalThis as any).getComputedStyle = csWithStyles
    restore()
  })

  it('pseudo-element with content:\'\' but transparent background is NOT extracted', () => {
    const { section } = setupSlide('<p>Content</p>')
    const p = section.querySelector('p')!
    mockRect(p, { left: 0, top: 0, width: 640, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {}],
    ])

    const csWithStyles = (globalThis as any).getComputedStyle
    ;(globalThis as any).getComputedStyle = (el: Element, pseudo?: string | null) => {
      if (pseudo === '::before' && el === section) {
        return {
          content: '""',
          backgroundColor: 'rgba(0, 0, 0, 0)',  // transparent
          position: 'absolute',
          top: '0px',
          left: '0px',
          width: '1280px',
          height: '12px',
        } as any
      }
      if (pseudo === '::after' && el === section) {
        return { content: 'none', backgroundColor: 'rgba(0,0,0,0)' } as any
      }
      return csWithStyles(el, pseudo)
    }

    const slides = extractSlides()
    const bar = slides[0].elements.find(
      (e: any) => e.type === 'container' && e.y === 0 && e.height === 12,
    )
    expect(bar).toBeUndefined()

    ;(globalThis as any).getComputedStyle = csWithStyles
    restore()
  })

  it('グローバル section::before (classless section と同色) — クラス付きスライドでも抑制', () => {
    // section::before { content:''; background: dark-navy } が全スライドに定義されているケース。
    // user class を持つスライド (cover など) でも誤ってバーを抽出しないことを確認。
    document.body.innerHTML = `
      <section id="regular" data-marpit-pagination="1"><p>Content</p></section>
      <section id="cover" data-marpit-pagination="2" class="cover"><h1>Title</h1></section>
    `
    const regularSection = document.getElementById('regular') as HTMLElement
    const coverSection = document.getElementById('cover') as HTMLElement
    const p = regularSection.querySelector('p')!
    const h1 = coverSection.querySelector('h1')!

    for (const s of [regularSection, coverSection]) {
      mockRect(s as Element, { left: 0, top: 0, width: 1280, height: 720 })
    }
    mockRect(p, { left: 58, top: 52, width: 1164, height: 24 })
    mockRect(h1, { left: 58, top: 52, width: 1164, height: 55 })

    const restore = mockStyles([
      [regularSection as Element, { backgroundColor: 'rgb(255,255,255)' }],
      [coverSection as Element, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {}],
      [h1, { fontSize: '46px', fontWeight: '700' }],
    ])

    const originalCS = (globalThis as any).getComputedStyle
    ;(globalThis as any).getComputedStyle = (el: Element, pseudo?: string | null) => {
      // global section::before: 全セクション同じ dark-navy
      if (pseudo === '::before' && (el === regularSection || el === coverSection)) {
        return {
          content: '""',
          backgroundColor: 'rgb(22, 50, 79)', // #16324f — dark navy
          position: 'absolute',
          top: '0px',
          left: '0px',
          width: '1280px',
          height: '16px',
          display: 'block',
        } as any
      }
      if (pseudo === '::after') {
        return { content: 'none', backgroundColor: 'rgba(0,0,0,0)' } as any
      }
      return originalCS(el, pseudo)
    }

    const slides = extractSlides()

    // 両スライドともバーが抽出されてはならない
    expect(slides).toHaveLength(2)
    for (const slide of slides) {
      const bar = slide.elements.find(
        (e: any) => e.type === 'container' && e.y === 0 && e.height === 16,
      )
      expect(bar).toBeUndefined()
    }

    ;(globalThis as any).getComputedStyle = originalCS
    restore()
  })

  it('クラス固有の decorator は classless section と異なる色なら引き続き抽出', () => {
    // section.decorated::before のみ青いバー。classless section の ::before は透明。
    // → decorated スライドのバーは抽出される。
    document.body.innerHTML = `
      <section id="regular" data-marpit-pagination="1"><p>Page 1</p></section>
      <section id="decorated" data-marpit-pagination="2" class="decorated"><p>Page 2</p></section>
    `
    const regularSection = document.getElementById('regular') as HTMLElement
    const decoratedSection = document.getElementById('decorated') as HTMLElement
    const p1 = regularSection.querySelector('p')!
    const p2 = decoratedSection.querySelector('p')!

    for (const s of [regularSection, decoratedSection]) {
      mockRect(s as Element, { left: 0, top: 0, width: 1280, height: 720 })
    }
    mockRect(p1, { left: 0, top: 50, width: 640, height: 24 })
    mockRect(p2, { left: 0, top: 50, width: 640, height: 24 })

    const restore = mockStyles([
      [regularSection as Element, { backgroundColor: 'rgb(255,255,255)' }],
      [decoratedSection as Element, { backgroundColor: 'rgb(255,255,255)' }],
      [p1, {}],
      [p2, {}],
    ])

    const originalCS = (globalThis as any).getComputedStyle
    ;(globalThis as any).getComputedStyle = (el: Element, pseudo?: string | null) => {
      if (pseudo === '::before' && el === decoratedSection) {
        return {
          content: '""',
          backgroundColor: 'rgb(37, 99, 235)', // class-specific blue (not on regular)
          position: 'absolute', top: '0px', left: '0px',
          width: '1280px', height: '12px', display: 'block',
        } as any
      }
      if (pseudo === '::before' && el === regularSection) {
        return { content: 'none', backgroundColor: 'rgba(0,0,0,0)' } as any
      }
      if (pseudo === '::after') {
        return { content: 'none', backgroundColor: 'rgba(0,0,0,0)' } as any
      }
      return originalCS(el, pseudo)
    }

    const slides = extractSlides()

    // regular スライド: バーなし
    expect(slides[0].elements.find((e: any) => e.type === 'container' && e.y === 0)).toBeUndefined()

    // decorated スライド: バーが抽出されている
    const decoratedBar = slides[1].elements.find(
      (e: any) => e.type === 'container' && e.y === 0 && e.height === 12,
    ) as any
    expect(decoratedBar).toBeDefined()
    expect(decoratedBar.style.backgroundColor).toBe('rgb(37, 99, 235)')

    ;(globalThis as any).getComputedStyle = originalCS
    restore()
  })
})

// ---------------------------------------------------------------------------
// extractTextRuns — breakLine deduplication: <br> followed by \n whitespace
// ---------------------------------------------------------------------------

describe('extractTextRuns — <br>+\\n does not produce double breakLine (via extractSlides)', () => {
  it('single breakLine between lines separated by <br>\\n', () => {
    const { section } = setupSlide('<p id="p1">Line A<br>\nLine B</p>')
    const pEl = section.querySelector('#p1')!
    mockRect(pEl, { left: 0, top: 100, width: 1280, height: 60 })
    const restore = mockStyles([[pEl, { display: 'block' }]])

    const slides = extractSlides()
    restore()

    const para = slides[0].elements.find((e: any) => e.type === 'paragraph')
    expect(para).toBeDefined()

    // There must be exactly ONE breakLine run between "Line A" and "Line B"
    const breaks = (para as any).runs.filter((r: any) => r.breakLine === true)
    expect(breaks).toHaveLength(1)
  })
})

// ---------------------------------------------------------------------------
// extractTableData — row background fallback from <tr> element
// ---------------------------------------------------------------------------

describe('extractTableData — tr background fallback (via extractSlides)', () => {
  it('applies tr background when td is transparent', () => {
    const { section } = setupSlide(`
      <table id="tbl">
        <tbody>
          <tr id="tr0"><td id="td0">Cell A</td></tr>
          <tr id="tr1"><td id="td1">Cell B</td></tr>
        </tbody>
      </table>
    `)
    const tbl = section.querySelector('#tbl')!
    const tr0 = section.querySelector('#tr0')!
    const tr1 = section.querySelector('#tr1')!
    const td0 = section.querySelector('#td0')!
    const td1 = section.querySelector('#td1')!
    mockRect(tbl, { left: 0, top: 100, width: 800, height: 80 })
    mockRect(tr0, { left: 0, top: 100, width: 800, height: 40 })
    mockRect(tr1, { left: 0, top: 140, width: 800, height: 40 })
    mockRect(td0, { left: 0, top: 100, width: 800, height: 40 })
    mockRect(td1, { left: 0, top: 140, width: 800, height: 40 })

    const restore = mockStyles([
      [tbl, { display: 'table', backgroundColor: 'rgba(0, 0, 0, 0)' }],
      [tr0, { display: 'table-row', backgroundColor: 'rgb(255, 255, 255)' }],
      [tr1, { display: 'table-row', backgroundColor: 'rgb(246, 248, 250)' }],
      [td0, { display: 'table-cell', backgroundColor: 'rgba(0, 0, 0, 0)' }],
      [td1, { display: 'table-cell', backgroundColor: 'rgba(0, 0, 0, 0)' }],
    ])

    const slides = extractSlides()
    restore()

    const table = slides[0].elements.find((e: any) => e.type === 'table') as any
    expect(table).toBeDefined()
    // Row 0: tr has white → cell backgroundColor = white
    expect(table!.rows[0].cells[0].style.backgroundColor).toBe(
      'rgb(255, 255, 255)',
    )
    // Row 1: tr has light gray → cell backgroundColor = light gray (from tr)
    expect(table!.rows[1].cells[0].style.backgroundColor).toBe(
      'rgb(246, 248, 250)',
    )
  })
})

// ---------------------------------------------------------------------------
// walkElements (<p>) — inline image beside text (Case A: no <br>)
// ---------------------------------------------------------------------------

describe('walkElements (<p>) — inline image beside text shifts paragraph x (via extractSlides)', () => {
  it('paragraph x shifts to image right edge and y to image baseline area', () => {
    const { section } = setupSlide('<p id="p1"><img id="img1"> inline text</p>')
    const pEl = section.querySelector('#p1')!
    const img = section.querySelector('#img1') as HTMLImageElement
    // Paragraph is taller than image: image = 300px, paragraph = 320px (image + text below)
    mockRect(pEl, { left: 79, top: 166, width: 1123, height: 320 })
    mockRect(img, { left: 79, top: 166, width: 300, height: 300 })
    // Mock naturalWidth/Height so it registers as a real image
    Object.defineProperty(img, 'naturalWidth', { value: 300, configurable: true })
    Object.defineProperty(img, 'naturalHeight', { value: 300, configurable: true })

    const restore = mockStyles([
      [pEl, { display: 'block' }],
      [img, { display: 'inline', filter: 'none', visibility: 'visible' }],
    ])

    const slides = extractSlides()
    restore()

    const para = slides[0].elements.find((e: any) => e.type === 'paragraph')
    expect(para).toBeDefined()
    // x must be after the image right edge (79 + 300 = 379)
    expect(para!.x).toBeCloseTo(379, 0)
    // Width reduced by image width (1123 - 300 = 823)
    expect(para!.width).toBeCloseTo(1123 - 300, 0)
    // y: CSS vertical-align:baseline aligns text with image bottom.
    // lineHeight defaults to 24px in test; inlineImgYOffset = max(0, 300-24) = 276
    // y = 166 + 276 = 442  (shifts down to near the image bottom area)
    expect(para!.y).toBeCloseTo(166 + 276, 0)
    // height = max(10, 320 - 276) = 44
    expect(para!.height).toBeCloseTo(44, 0)
  })
})

// ---------------------------------------------------------------------------
// walkElements (<p>) — inline image above text (Case B: <br> after image)
// ---------------------------------------------------------------------------

describe('walkElements (<p>) — inline image above text shifts paragraph y (via extractSlides)', () => {
  it('paragraph y shifts to image bottom when <img><br>text pattern is used', () => {
    const { section } = setupSlide('<p id="p1"><img id="img1"><br>Caption text</p>')
    const pEl = section.querySelector('#p1')!
    const img = section.querySelector('#img1') as HTMLImageElement
    mockRect(pEl, { left: 79, top: 226, width: 1123, height: 354 })
    mockRect(img, { left: 79, top: 226, width: 300, height: 300 })
    Object.defineProperty(img, 'naturalWidth', { value: 300, configurable: true })
    Object.defineProperty(img, 'naturalHeight', { value: 300, configurable: true })

    const restore = mockStyles([
      [pEl, { display: 'block' }],
      [img, { display: 'inline', filter: 'none', visibility: 'visible' }],
    ])

    const slides = extractSlides()
    restore()

    const para = slides[0].elements.find((e: any) => e.type === 'paragraph')
    expect(para).toBeDefined()
    // Paragraph y must be at image bottom (226 + 300 = 526)
    expect(para!.y).toBeCloseTo(526, 0)
    // Height reduced by image height (354 - 300 = 54)
    expect(para!.height).toBeCloseTo(54, 0)
    // x unchanged
    expect(para!.x).toBeCloseTo(79, 0)
  })
})

// -----------------------------------------------------------------------
// Regression: flex/grid container — direct text nodes lost when a child
// element also produces blockChildren.
//
// Scenario: table-of-contents slide (目次ページ) pattern.
//   <div class="agenda-wrap">   (display:grid)
//     <div class="agenda-item"> (display:flex, has both a badge span AND text)
//       <span class="agenda-num">1</span>
//       Background and purpose
//     </div>
//     ...
//   </div>
//
// Bug: walkElements(agenda-item) → finds span (blockChild) → blockChildren.length > 0
//   → emits container{children:[para("1")]} and drops the direct text node
//   "Background and purpose".
//
// Fix: when a flex/grid container has block-level children AND direct text
//   nodes that are non-empty, those text nodes must also be captured.
// -----------------------------------------------------------------------

describe('flex/grid container: direct text nodes preserved when child produces blockChildren', () => {
  it('text node sibling to badge span inside flex item is NOT lost', () => {
    // Represents: <div class="agenda-item"><span class="num">1</span> Agenda text</div>
    // Both the badge span AND the plain text must appear in output.
    const { section } = setupSlide(`
      <div id="agenda-wrap">
        <div id="agenda-item">
          <span id="agenda-num">1</span>
          Agenda text
        </div>
      </div>
    `)
    const wrap = section.querySelector('#agenda-wrap')!
    const item = section.querySelector('#agenda-item')!
    const num = section.querySelector('#agenda-num')!

    mockRect(wrap, { left: 58, top: 100, width: 1164, height: 200 })
    mockRect(item, { left: 58, top: 100, width: 582, height: 40 })
    mockRect(num, { left: 58, top: 104, width: 28, height: 28 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [wrap, { display: 'grid', gridTemplateColumns: '1fr 1fr' }],
      [
        item,
        {
          display: 'flex',
          alignItems: 'flex-start',
          gap: '14px',
          color: 'rgb(0,0,0)',
          fontSize: '20px',
          fontFamily: 'Arial',
          fontWeight: '400',
          lineHeight: '30px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [
        num,
        {
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: 'rgb(15, 108, 189)',
          color: 'rgb(255,255,255)',
          borderRadius: '50%',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          lineHeight: '28px',
          textAlign: 'left',
        },
      ],
    ])

    const slides = extractSlides()

    // Collect all text from all elements recursively
    function collectTexts(els: any[]): string[] {
      const texts: string[] = []
      for (const el of els) {
        if (el.runs) {
          for (const r of el.runs) {
            if (!r.breakLine && r.text.trim()) texts.push(r.text.trim())
          }
        }
        if (el.children) texts.push(...collectTexts(el.children))
        if (el.items) {
          for (const item of el.items) {
            for (const r of item.runs ?? []) {
              if (!r.breakLine && r.text.trim()) texts.push(r.text.trim())
            }
          }
        }
      }
      return texts
    }

    const allTexts = collectTexts(slides[0].elements)

    // Badge number "1" must appear
    expect(allTexts).toContain('1')
    // The direct text node "Agenda text" MUST NOT be lost
    expect(allTexts.some((t) => t.includes('Agenda text'))).toBe(true)

    function findParagraphWithText(els: any[], text: string): any {
      for (const el of els) {
        if (
          el.type === 'paragraph' &&
          el.runs?.some((r: any) => r.text?.includes(text))
        )
          return el
        if (el.children) {
          const found = findParagraphWithText(el.children, text)
          if (found) return found
        }
      }
      return null
    }

    const agendaPara = findParagraphWithText(slides[0].elements, 'Agenda text')
    expect(agendaPara).toBeDefined()
    expect(agendaPara.x).toBe(100)

    restore()
  })

  it('flex item with left padding still offsets recovered text after badge and gap', () => {
    const { section } = setupSlide(`
      <div id="agenda-wrap">
        <div id="agenda-item">
          <span id="agenda-num">1</span>
          Agenda text
        </div>
      </div>
    `)
    const wrap = section.querySelector('#agenda-wrap')!
    const item = section.querySelector('#agenda-item')!
    const num = section.querySelector('#agenda-num')!

    mockRect(wrap, { left: 58, top: 100, width: 1164, height: 200 })
    mockRect(item, { left: 58, top: 100, width: 582, height: 40 })
    mockRect(num, { left: 74, top: 104, width: 28, height: 28 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [wrap, { display: 'grid', gridTemplateColumns: '1fr 1fr' }],
      [
        item,
        {
          display: 'flex',
          alignItems: 'flex-start',
          gap: '14px',
          paddingLeft: '16px',
          color: 'rgb(0,0,0)',
          fontSize: '20px',
          fontFamily: 'Arial',
          fontWeight: '400',
          lineHeight: '30px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [
        num,
        {
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: 'rgb(15, 108, 189)',
          color: 'rgb(255,255,255)',
          borderRadius: '50%',
          fontSize: '14px',
          fontFamily: 'Arial',
          fontWeight: '700',
          lineHeight: '28px',
          textAlign: 'left',
        },
      ],
    ])

    const slides = extractSlides()

    function findParagraphWithText(els: any[], text: string): any {
      for (const el of els) {
        if (
          el.type === 'paragraph' &&
          el.runs?.some((r: any) => r.text?.includes(text))
        ) {
          return el
        }
        if (el.children) {
          const found = findParagraphWithText(el.children, text)
          if (found) return found
        }
      }
      return null
    }

    const agendaPara = findParagraphWithText(slides[0].elements, 'Agenda text')
    expect(agendaPara).toBeDefined()
    expect(agendaPara.x).toBe(116)

    restore()
  })

  it('two-column grid with badge+text items — all item texts extracted', () => {
    // Minimal reproduction of a 2-item agenda grid where each cell has
    // a numbered badge and a description text node.
    const { section } = setupSlide(`
      <div id="grid">
        <div id="item1"><span id="n1">1</span> Topic Alpha</div>
        <div id="item2"><span id="n2">2</span> Topic Beta</div>
      </div>
    `)
    const grid = section.querySelector('#grid')!
    const item1 = section.querySelector('#item1')!
    const n1 = section.querySelector('#n1')!
    const item2 = section.querySelector('#item2')!
    const n2 = section.querySelector('#n2')!

    mockRect(grid, { left: 58, top: 100, width: 1164, height: 80 })
    mockRect(item1, { left: 58, top: 100, width: 582, height: 40 })
    mockRect(n1, { left: 58, top: 106, width: 28, height: 28 })
    mockRect(item2, { left: 640, top: 100, width: 582, height: 40 })
    mockRect(n2, { left: 640, top: 106, width: 28, height: 28 })

    const badgeStyle = {
      display: 'inline-flex',
      alignItems: 'center',
      justifyContent: 'center',
      backgroundColor: 'rgb(15, 108, 189)',
      color: 'rgb(255,255,255)',
      borderRadius: '50%',
      fontSize: '14px',
      fontFamily: 'Arial',
      fontWeight: '700',
      lineHeight: '28px',
      textAlign: 'left',
    }
    const itemStyle = {
      display: 'flex',
      alignItems: 'flex-start',
      gap: '14px',
      color: 'rgb(0,0,0)',
      fontSize: '20px',
      fontFamily: 'Arial',
      fontWeight: '400',
      lineHeight: '30px',
      textAlign: 'left',
      backgroundColor: 'rgba(0,0,0,0)',
    }

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [grid, { display: 'grid', gridTemplateColumns: '1fr 1fr' }],
      [item1, itemStyle],
      [n1, badgeStyle],
      [item2, itemStyle],
      [n2, badgeStyle],
    ])

    const slides = extractSlides()

    function collectTexts(els: any[]): string[] {
      const texts: string[] = []
      for (const el of els) {
        if (el.runs) {
          for (const r of el.runs) {
            if (!r.breakLine && r.text.trim()) texts.push(r.text.trim())
          }
        }
        if (el.children) texts.push(...collectTexts(el.children))
      }
      return texts
    }

    const allTexts = collectTexts(slides[0].elements)

    // Both descriptions must be present
    expect(allTexts.some((t) => t.includes('Topic Alpha'))).toBe(true)
    expect(allTexts.some((t) => t.includes('Topic Beta'))).toBe(true)

    function findParagraphWithText(els: any[], text: string): any {
      for (const el of els) {
        if (
          el.type === 'paragraph' &&
          el.runs?.some((r: any) => r.text?.includes(text))
        ) {
          return el
        }
        if (el.children) {
          const found = findParagraphWithText(el.children, text)
          if (found) return found
        }
      }
      return null
    }

    const alphaPara = findParagraphWithText(slides[0].elements, 'Topic Alpha')
    const betaPara = findParagraphWithText(slides[0].elements, 'Topic Beta')

    expect(alphaPara).toBeDefined()
    expect(betaPara).toBeDefined()
    expect(alphaPara.x).toBe(100)
    expect(betaPara.x).toBe(682)

    restore()
  })

  it('block container with block child + inline child + text tail preserves inline and tail text', () => {
    const { section } = setupSlide(`
      <div id="wrap"><p id="block">Block child</p><span id="inline">Inline part</span> tail text</div>
    `)
    const wrap = section.querySelector('#wrap')!
    const block = section.querySelector('#block')!
    const inline = section.querySelector('#inline')!

    mockRect(wrap, { left: 40, top: 90, width: 700, height: 80 })
    mockRect(block, { left: 40, top: 90, width: 300, height: 28 })
    mockRect(inline, { left: 40, top: 130, width: 120, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        wrap,
        {
          display: 'block',
          color: 'rgb(0,0,0)',
          fontSize: '18px',
          fontFamily: 'Arial',
          fontWeight: '400',
          lineHeight: '28px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [block, { display: 'block' }],
      [
        inline,
        {
          display: 'inline',
          color: 'rgb(0,0,0)',
          fontSize: '18px',
          fontFamily: 'Arial',
          fontWeight: '700',
          lineHeight: '28px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
    ])

    const slides = extractSlides()

    function collectTexts(els: any[]): string[] {
      const texts: string[] = []
      for (const el of els) {
        if (el.runs) {
          for (const r of el.runs) {
            if (!r.breakLine && r.text.trim()) texts.push(r.text.trim())
          }
        }
        if (el.children) texts.push(...collectTexts(el.children))
      }
      return texts
    }

    const allTexts = collectTexts(slides[0].elements)
    expect(allTexts.some((t) => t.includes('Block child'))).toBe(true)
    expect(allTexts.some((t) => t.includes('Inline part'))).toBe(true)
    // Direct text nodes are intentionally NOT recovered for block containers
    // (only flex/grid containers recover direct text nodes — see mermaid regression).
    // 'tail text' is a direct TEXT_NODE and will not appear in the PPTX.
    expect(allTexts.some((t) => t.includes('tail text'))).toBe(false)

    function findParagraphWithText(els: any[], text: string): any {
      for (const el of els) {
        if (
          el.type === 'paragraph' &&
          el.runs?.some((r: any) => r.text?.includes(text))
        ) {
          return el
        }
        if (el.children) {
          const found = findParagraphWithText(el.children, text)
          if (found) return found
        }
      }
      return null
    }

    const inlinePara = findParagraphWithText(slides[0].elements, 'Inline part')
    expect(inlinePara).toBeDefined()
    expect(inlinePara.x).toBe(40)

    restore()
  })
})

// -----------------------------------------------------------------------
// Regression: mermaid SVG raw source text must not leak into PPTX.
//
// When mermaid.js processes a <div class="mermaid"> element it replaces
// the element's innerHTML with an <svg> element.  In some cases (CDN latency,
// partial rendering) the original text nodes (diagram syntax like
// "flowchart LR\n  A[X] --> B[Y]") can still be present in the DOM at the
// time the DOM walker runs.
//
// Bug (ADR-15 regression): the shallow text recovery added for flex/grid
// containers was inadvertently applied to ALL containers with block children.
// TEXT_NODEs in a block container whose only block child is an SVG image
// would be picked up and emitted as a text paragraph on top of the image.
//
// Fix: restrict TEXT_NODE recovery to flex/grid containers only.
// -----------------------------------------------------------------------
describe('mermaid: raw source text nodes must not appear in PPTX output', () => {
  it('block container with rendered SVG child does not capture orphaned text node', () => {
    // Simulates a <div class="mermaid"> whose rendering was only partially
    // complete: the SVG was inserted but the original text node (diagram
    // source syntax) was not yet removed.
    const { section } = setupSlide(`
      <div id="mermaid-div">
        flowchart LR
          A[Source] --&gt; B[Result]
        <svg id="mermaid-svg"></svg>
      </div>
    `)
    const mermaidDiv = section.querySelector('#mermaid-div')!
    const mermaidSvg = section.querySelector('#mermaid-svg')!

    mockRect(mermaidDiv, { left: 100, top: 50, width: 800, height: 400 })
    mockRect(mermaidSvg, { left: 100, top: 50, width: 800, height: 400 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        mermaidDiv,
        {
          display: 'block',
          color: 'rgb(0,0,0)',
          fontSize: '16px',
          fontFamily: 'Arial',
          fontWeight: '400',
          lineHeight: '24px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [mermaidSvg, { display: 'block' }],
    ])

    const slides = extractSlides()

    function collectAllTexts(els: any[]): string[] {
      const texts: string[] = []
      for (const el of els) {
        if (el.runs) {
          for (const r of el.runs) {
            if (!r.breakLine && r.text.trim()) texts.push(r.text.trim())
          }
        }
        if (el.children) texts.push(...collectAllTexts(el.children))
      }
      return texts
    }

    const allTexts = collectAllTexts(slides[0].elements)

    // The raw mermaid source syntax MUST NOT appear as a text element.
    expect(allTexts.some((t) => t.includes('flowchart'))).toBe(false)
    expect(allTexts.some((t) => t.includes('-->'))).toBe(false)
    expect(allTexts.some((t) => t.includes('Source'))).toBe(false)

    // The SVG should be captured as an image element.
    function hasImageElement(els: any[]): boolean {
      for (const el of els) {
        if (el.type === 'image') return true
        if (el.children && hasImageElement(el.children)) return true
      }
      return false
    }
    expect(hasImageElement(slides[0].elements)).toBe(true)

    restore()
  })

  it('flex container with rendered SVG child still recovers sibling text node', () => {
    // A flex container that has both an SVG child AND a direct text node
    // (e.g., an icon+label layout) SHOULD still recover the text node.
    const { section } = setupSlide(`
      <div id="icon-label">
        <svg id="icon-svg"></svg>
        Label text
      </div>
    `)
    const iconLabel = section.querySelector('#icon-label')!
    const iconSvg = section.querySelector('#icon-svg')!

    mockRect(iconLabel, { left: 50, top: 100, width: 300, height: 40 })
    mockRect(iconSvg, { left: 50, top: 106, width: 28, height: 28 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [
        iconLabel,
        {
          display: 'flex',
          alignItems: 'center',
          gap: '8px',
          color: 'rgb(0,0,0)',
          fontSize: '18px',
          fontFamily: 'Arial',
          fontWeight: '400',
          lineHeight: '28px',
          textAlign: 'left',
          backgroundColor: 'rgba(0,0,0,0)',
        },
      ],
      [iconSvg, { display: 'block' }],
    ])

    const slides = extractSlides()

    function collectAllTexts(els: any[]): string[] {
      const texts: string[] = []
      for (const el of els) {
        if (el.runs) {
          for (const r of el.runs) {
            if (!r.breakLine && r.text.trim()) texts.push(r.text.trim())
          }
        }
        if (el.children) texts.push(...collectAllTexts(el.children))
      }
      return texts
    }

    const allTexts = collectAllTexts(slides[0].elements)

    // In a flex container, the direct text node 'Label text' MUST be recovered.
    expect(allTexts.some((t) => t.includes('Label text'))).toBe(true)

    restore()
  })

  it('extractNestedImages skips emoji img with only alt-text Extended_Pictographic cue', () => {
    // The Twemoji CDN URL always contains "twemoji", but isEmojiImg() also
    // accepts imgs whose alt text is a single Extended Pictographic character.
    // extractNestedImages must use the full isEmojiImg() check so an emoji img
    // is never extracted as a floating image shape when its src is non-standard.
    const { section } = setupSlide(`
      <p id="p-emoji">
        Verify operation
        <img id="img-emoji" alt="✅" />
      </p>
    `)
    const p = section.querySelector('#p-emoji')!
    const img = section.querySelector('#img-emoji') as HTMLImageElement

    mockRect(p, { left: 79, top: 100, width: 400, height: 30 })
    mockRect(img, { left: 420, top: 103, width: 16, height: 16 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, { display: 'block', color: 'rgb(0,0,0)', fontSize: '16px', fontFamily: 'Arial', fontWeight: '400', lineHeight: '24px', textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)' }],
      [img, { display: 'inline' }],
    ])

    const slides = extractSlides()

    // The emoji img should NOT appear as a type:'image' element anywhere in the tree.
    function hasImageType(els: any[]): boolean {
      for (const el of els) {
        if (el.type === 'image') return true
        if (el.children && hasImageType(el.children)) return true
      }
      return false
    }
    expect(hasImageType(slides[0].elements)).toBe(false)

    restore()
  })
})

// -----------------------------------------------------------------------
// emoji テキストを含む inline-only flex 子要素の幅拡張
// -----------------------------------------------------------------------

describe('flex child with emoji text — width extended to parent right edge', () => {
  it('span containing Twemoji alt text inside flex row gets width extended to row right edge', () => {
    // Reproduces slide 18: <div flex-row><span badge>3</span><span>Verify operation ✅</span></div>
    // The text span's intrinsic width is just enough to fit the text, but PPTX
    // font rendering may make ✅ slightly wider than the Twemoji 1em image → wraps.
    const { section } = setupSlide(`
      <div id="flex-row">
        <span id="badge">3</span>
        <span id="text-span">Verify operation <img id="emoji-img" class="emoji" alt="✅" src="https://twemoji/2705.svg"></span>
      </div>
    `)

    const flexRow = section.querySelector('#flex-row')! as HTMLElement
    const badge = section.querySelector('#badge')! as HTMLElement
    const textSpan = section.querySelector('#text-span')! as HTMLElement
    const emojiImg = section.querySelector('#emoji-img')! as HTMLImageElement

    // flex row spans full slide width (0..1280), text span starts at x=78
    mockRect(section, { left: 0, top: 0, width: 1280, height: 720 })
    mockRect(flexRow, { left: 40, top: 200, width: 900, height: 28 })
    mockRect(badge, { left: 40, top: 200, width: 28, height: 28 })
    mockRect(textSpan, { left: 78, top: 200, width: 180, height: 28 })
    mockRect(emojiImg, { left: 238, top: 203, width: 18, height: 18 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [flexRow, { display: 'flex', alignItems: 'center', gap: '10px', backgroundColor: 'rgba(0,0,0,0)' }],
      [badge, { display: 'inline-flex', backgroundColor: 'rgb(0,102,204)', color: 'rgb(255,255,255)', fontSize: '14px', fontFamily: 'Arial', fontWeight: '400' }],
      [textSpan, { display: 'inline', color: 'rgb(0,0,0)', fontSize: '16px', fontFamily: 'Arial', fontWeight: '400', backgroundColor: 'rgba(0,0,0,0)', lineHeight: '24px' }],
      [emojiImg, { display: 'inline' }],
    ])

    const slides = extractSlides()
    restore()

    // Find all paragraph elements recursively
    function findParagraphs(els: any[]): any[] {
      const result: any[] = []
      for (const el of els) {
        if (el.type === 'paragraph') result.push(el)
        if (el.children) result.push(...findParagraphs(el.children))
      }
      return result
    }
    const paras = findParagraphs(slides[0].elements)

    // The paragraph containing "Verify operation" + "✅" should have its width
    // extended to the flex row's right edge (40 + 900 = 940; minus span x=78 → 862)
    const verifyPara = paras.find((p: any) =>
      p.runs?.some((r: any) => r.text?.includes('Verify operation')),
    )
    expect(verifyPara).toBeDefined()
    // Width should be extended beyond the intrinsic 180px to accommodate PPTX emoji rendering
    expect(verifyPara.width).toBeGreaterThan(180)
    // Width should be approximately flexRow.right - textSpan.left = (40+900) - 78 = 862
    expect(verifyPara.width).toBeCloseTo(862, 0)
  })
})

// -----------------------------------------------------------------------
// leading + mid-line badges in <p> — slide 34 regression
// badge-only paragraph → all badges as shapes (no surrounding text)
// mixed paragraph (text + badges) → leading badge as shape, non-leading as inline highlight
// -----------------------------------------------------------------------

describe('all badges in <p> emitted as shapes', () => {
  it('leading badge → shape, mid-line badge → bg-only shape + text without bg when paragraph has surrounding text', () => {
    // <p><span badge>1</span> Install <span badge>2</span> Step two</p>
    // b1 is leading (left=50=para.left) → container shape
    // b2 is mid-line (left=200 >> para.left+8) → inline highlight in paragraph
    // "Install" TEXT NODE → containerHasNonBadgeText = true → leading filter applied
    const { section } = setupSlide(`
      <p id="para">
        <span id="b1">1</span> Install
        <span id="b2">2</span> Step two
      </p>
    `)
    const para = section.querySelector('#para')! as HTMLElement
    const b1 = section.querySelector('#b1')! as HTMLElement
    const b2 = section.querySelector('#b2')! as HTMLElement

    mockRect(para, { left: 50, top: 100, width: 600, height: 36 })
    mockRect(b1,   { left: 50, top: 104, width: 28, height: 28 })   // leading (x=para.left)
    mockRect(b2,   { left: 200, top: 104, width: 28, height: 28 })  // mid-line

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [para, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [b1, {
        display: 'inline-flex', backgroundColor: 'rgb(0,102,204)',
        color: 'rgb(255,255,255)', borderRadius: '999px',
        fontSize: '14px', fontFamily: 'Arial', fontWeight: '700',
      }],
      [b2, {
        display: 'inline-flex', backgroundColor: 'rgb(0,102,204)',
        color: 'rgb(255,255,255)', borderRadius: '999px',
        fontSize: '14px', fontFamily: 'Arial', fontWeight: '700',
      }],
    ])

    const slides = extractSlides()
    restore()
    const elements = slides[0].elements

    // Leading badge b1 → container shape (rounded corners in PPTX)
    const containers = elements.filter((e: any) => e.type === 'container')
    expect(containers).toHaveLength(2)  // leading (with runs) + mid-line (bg-only)
    expect((containers[0] as any).style.backgroundColor).toBe('rgb(0,102,204)')

    // Paragraph MUST exist with surrounding text
    const paragraph = elements.find((e: any) => e.type === 'paragraph') as any
    expect(paragraph).toBeDefined()

    // b1 text must NOT appear in paragraph runs (it's in the shape)
    const run1 = paragraph.runs?.find((r: any) => r.text === '1')
    expect(run1).toBeUndefined()

    // b2 (non-leading) text MUST appear in paragraph runs WITHOUT backgroundColor
    // (bg-only container shape provides the visual background)
    const run2 = paragraph.runs?.find((r: any) => r.text === '2')
    expect(run2).toBeDefined()
    expect(run2.backgroundColor).toBeUndefined()

    // Surrounding text still present
    const installRun = paragraph.runs?.find((r: any) => r.text?.includes('Install'))
    expect(installRun).toBeDefined()
    const stepRun = paragraph.runs?.find((r: any) => r.text?.includes('Step two'))
    expect(stepRun).toBeDefined()
  })

  it('badge-only paragraph: ALL badges extracted as shapes regardless of position', () => {
    // <p><span>HIGH</span><span>MED</span><span>LOW</span></p>
    // No surrounding text → containerHasNonBadgeText = false
    // → leading filter NOT applied → all 3 badges become shapes
    const { section } = setupSlide(`
      <p id="para"><span id="b1">HIGH</span><span id="b2">MED</span><span id="b3">LOW</span></p>
    `)
    const para = section.querySelector('#para')! as HTMLElement
    const b1 = section.querySelector('#b1')! as HTMLElement
    const b2 = section.querySelector('#b2')! as HTMLElement
    const b3 = section.querySelector('#b3')! as HTMLElement

    mockRect(para, { left: 50, top: 200, width: 400, height: 40 })
    mockRect(b1,   { left: 50,  top: 204, width: 80, height: 32 })  // leading
    mockRect(b2,   { left: 138, top: 204, width: 80, height: 32 })  // non-leading
    mockRect(b3,   { left: 226, top: 204, width: 80, height: 32 })  // non-leading

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [para, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [b1, {
        display: 'inline-flex', backgroundColor: 'rgb(192,86,33)',
        color: 'rgb(255,255,255)', borderRadius: '16px',
        fontSize: '14px', fontFamily: 'Arial', fontWeight: '700',
      }],
      [b2, {
        display: 'inline-flex', backgroundColor: 'rgb(221,107,32)',
        color: 'rgb(255,255,255)', borderRadius: '16px',
        fontSize: '14px', fontFamily: 'Arial', fontWeight: '700',
      }],
      [b3, {
        display: 'inline-flex', backgroundColor: 'rgb(47,133,90)',
        color: 'rgb(255,255,255)', borderRadius: '16px',
        fontSize: '14px', fontFamily: 'Arial', fontWeight: '700',
      }],
    ])

    const slides = extractSlides()
    restore()
    const elements = slides[0].elements

    // All 3 badges → container shapes (no leading filter for badge-only paragraph)
    const containers = elements.filter((e: any) => e.type === 'container')
    expect(containers).toHaveLength(3)
    expect((containers[0] as any).style.backgroundColor).toBe('rgb(192,86,33)')
    expect((containers[1] as any).style.backgroundColor).toBe('rgb(221,107,32)')
    expect((containers[2] as any).style.backgroundColor).toBe('rgb(47,133,90)')

    // No paragraph needed (all content is in shapes)
    const paragraph = elements.find((e: any) => e.type === 'paragraph')
    expect(paragraph).toBeUndefined()

    restore()
  })
})


// -----------------------------------------------------------------------
// image embedded between list items (no blank lines — image inside <li>)
// -----------------------------------------------------------------------

describe('image embedded between list items (via extractSlides)', () => {
  it('splits list around image-containing <li> so image preserves position', () => {
    const { section } = setupSlide(`
      <ul id="ul">
        <li id="li1">First item<br><img id="img" src="img.png" width="200" height="100"></li>
        <li id="li2">Second item</li>
      </ul>
    `)
    const ul = section.querySelector('#ul')!
    const li1 = section.querySelector('#li1')!
    const li2 = section.querySelector('#li2')!
    const img = section.querySelector('#img') as HTMLImageElement

    Object.defineProperty(img, 'naturalWidth', { value: 200, configurable: true })
    Object.defineProperty(img, 'naturalHeight', { value: 100, configurable: true })

    mockRect(ul, { left: 78, top: 100, width: 600, height: 190 })
    mockRect(li1, { left: 78, top: 100, width: 600, height: 130 })
    mockRect(img, { left: 78, top: 130, width: 200, height: 100 })
    mockRect(li2, { left: 78, top: 230, width: 600, height: 30 })

    const liStyle = {
      display: 'list-item',
      fontSize: '16px',
      fontFamily: 'Arial',
      color: 'rgb(0,0,0)',
      fontWeight: '400',
      fontStyle: 'normal',
      textAlign: 'left',
      lineHeight: '24px',
    }
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ul, { display: 'block', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px' }],
      [li1, liStyle],
      [li2, liStyle],
      [img, { display: 'inline', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px', backgroundColor: 'rgba(0,0,0,0)' }],
    ])

    const slides = extractSlides()
    const els = slides[0].elements

    const lists = els.filter((e: any) => e.type === 'list') as any[]
    const images = els.filter((e: any) => e.type === 'image') as any[]

    expect(lists).toHaveLength(2)
    expect(images).toHaveLength(1)

    expect(lists[0].items[0]).toMatchObject({ text: 'First item' })
    expect(lists[1].items[0]).toMatchObject({ text: 'Second item' })

    expect(images[0].y).toBeCloseTo(130)
    expect(lists[1].y).toBeGreaterThanOrEqual(images[0].y + images[0].height - 5)

    restore()
  })

  it('image-only <li> emits standalone image without empty list items', () => {
    const { section } = setupSlide(`
      <ul id="ul">
        <li id="li1">Before</li>
        <li id="li2"><img id="img" src="img.png" width="100" height="80"></li>
        <li id="li3">After</li>
      </ul>
    `)
    const ul = section.querySelector('#ul')!
    const li1 = section.querySelector('#li1')!
    const li2 = section.querySelector('#li2')!
    const li3 = section.querySelector('#li3')!
    const img = section.querySelector('#img') as HTMLImageElement

    Object.defineProperty(img, 'naturalWidth', { value: 100, configurable: true })
    Object.defineProperty(img, 'naturalHeight', { value: 80, configurable: true })

    mockRect(ul, { left: 78, top: 100, width: 600, height: 160 })
    mockRect(li1, { left: 78, top: 100, width: 600, height: 30 })
    mockRect(li2, { left: 78, top: 130, width: 600, height: 80 })
    mockRect(img, { left: 78, top: 130, width: 100, height: 80 })
    mockRect(li3, { left: 78, top: 210, width: 600, height: 30 })

    const liStyle = {
      display: 'list-item',
      fontSize: '16px',
      fontFamily: 'Arial',
      color: 'rgb(0,0,0)',
      fontWeight: '400',
      fontStyle: 'normal',
      textAlign: 'left',
      lineHeight: '24px',
    }
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ul, { display: 'block', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px' }],
      [li1, liStyle],
      [li2, liStyle],
      [li3, liStyle],
      [img, { display: 'inline', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px', backgroundColor: 'rgba(0,0,0,0)' }],
    ])

    const slides = extractSlides()
    const els = slides[0].elements

    const images = els.filter((e: any) => e.type === 'image') as any[]
    const allItems = els.filter((e: any) => e.type === 'list').flatMap((l: any) => l.items)

    expect(images).toHaveLength(1)
    expect(allItems.some((i: any) => i.text === '')).toBe(false)
    expect(allItems[0]).toMatchObject({ text: 'Before' })
    expect(allItems[allItems.length - 1]).toMatchObject({ text: 'After' })

    restore()
  })
})

// ===========================================================================
// Loose list — <li><p>…</p><p>…</p></li> paragraph separator
// ===========================================================================
// When Markdown has blank lines between continuation lines of the same item,
// markdown-it wraps each paragraph in <p>.  In HTML the <p> elements have
// natural margin spacing.  In PPTX each paragraph must become a separate line
// via a breakLine run — otherwise both texts appear merged on one line.
// ===========================================================================
describe('loose list — multiple <p> inside single <li>', () => {
  function buildLooseList() {
    const { section } = setupSlide(`
      <ul id="ul">
        <li id="li1">
          <p id="p1">Paragraph A</p>
          <p id="p2">Paragraph B</p>
        </li>
        <li id="li2"><p id="p3">Second item</p></li>
      </ul>
    `)
    const ul = section.querySelector('#ul')!
    const li1 = section.querySelector('#li1')!
    const li2 = section.querySelector('#li2')!
    const p1 = section.querySelector('#p1')!
    const p2 = section.querySelector('#p2')!
    const p3 = section.querySelector('#p3')!

    mockRect(ul, { left: 50, top: 100, width: 600, height: 120 })
    mockRect(li1, { left: 60, top: 100, width: 580, height: 80 })
    mockRect(li2, { left: 60, top: 180, width: 580, height: 40 })
    mockRect(p1, { left: 70, top: 100, width: 560, height: 24 })
    mockRect(p2, { left: 70, top: 130, width: 560, height: 24 })
    mockRect(p3, { left: 70, top: 180, width: 560, height: 24 })

    const liStyle = {
      display: 'list-item', fontSize: '16px', fontFamily: 'Arial',
      color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal',
      textAlign: 'left', lineHeight: '24px',
      backgroundColor: 'rgba(0,0,0,0)',
    }
    const pStyle = {
      display: 'block', fontSize: '16px', fontFamily: 'Arial',
      color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal',
      textAlign: 'left', lineHeight: '24px',
      backgroundColor: 'rgba(0,0,0,0)',
    }
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ul, { display: 'block', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px', backgroundColor: 'rgba(0,0,0,0)' }],
      [li1, liStyle], [li2, liStyle],
      [p1, pStyle], [p2, pStyle], [p3, pStyle],
    ])

    return { restore }
  }

  it('<p> per paragraph in loose <li> produces a breakLine run between them', () => {
    const { restore } = buildLooseList()
    const slides = extractSlides()
    const listEl = slides[0].elements.find((e: any) => e.type === 'list') as any
    const item = listEl.items[0] // first <li> with two <p>

    const runs = item.runs
    const breakIdx = runs.findIndex((r: any) => r.breakLine === true)

    expect(breakIdx).not.toBe(-1) // must have a break between paragraphs
    expect(runs[0].text).toBe('Paragraph A')
    expect(runs[breakIdx + 1].text).toBe('Paragraph B')

    restore()
  })

  it('second item (single <p>) has no spurious breakLine', () => {
    const { restore } = buildLooseList()
    const slides = extractSlides()
    const listEl = slides[0].elements.find((e: any) => e.type === 'list') as any
    const second = listEl.items[1] // second <li>

    expect(second.runs.some((r: any) => r.breakLine)).toBe(false)

    restore()
  })
})

// ===========================================================================
// Pipeline integration: HTML → extractSlides → toListTextProps
// ===========================================================================
// These tests traverse the FULL pipeline so that regressions in EITHER
// dom-walker OR slide-builder are caught in a single place.
//
// Rationale: Today's bug was that dom-walker correctly produced breakLine runs
// AND slide-builder correctly mapped them to TextProps — but a previous version
// of slide-builder silently dropped the marL alignment.  Tests for individual
// layers cannot catch cross-layer interactions like this.
//
// Each test here starts from raw HTML → extractSlides() → toListTextProps()
// and asserts on the final PptxGenJS TextProps structure.
// ===========================================================================
describe('pipeline integration — HTML to toListTextProps', () => {
  it('<li>Line one<br>Line two</li> → continuation paragraph uses invisible bullet (marL correct)', () => {
    // This test guards the FULL pipeline from HTML to PptxGenJS TextProps.
    // The invisible bullet (characterCode:'200B') is how we achieve correct
    // marL in the continuation paragraph — matching HTML text-start position.
    const { section } = setupSlide(`
      <ul id="ul">
        <li id="li1">Line one<br>Line two</li>
        <li id="li2">Normal item</li>
      </ul>
    `)
    const ul = section.querySelector('#ul')!
    const li1 = section.querySelector('#li1')!
    const li2 = section.querySelector('#li2')!

    mockRect(ul, { left: 50, top: 100, width: 600, height: 80 })
    mockRect(li1, { left: 60, top: 100, width: 580, height: 40 })
    mockRect(li2, { left: 60, top: 140, width: 580, height: 40 })

    const liStyle = {
      display: 'list-item', fontSize: '16px', fontFamily: 'Arial',
      color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal',
      textAlign: 'left', lineHeight: '24px',
      backgroundColor: 'rgba(0,0,0,0)',
    }
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ul, { display: 'block', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px', backgroundColor: 'rgba(0,0,0,0)' }],
      [li1, liStyle], [li2, liStyle],
    ])

    const slides = extractSlides()
    const listEl = slides[0].elements.find((e: any) => e.type === 'list') as any
    const firstItem = listEl.items[0]

    // 1. dom-walker layer: item must have breakLine runs
    const breakIdx = firstItem.runs.findIndex((r: any) => r.breakLine === true)
    expect(breakIdx).not.toBe(-1)
    expect(firstItem.runs[0].text).toBe('Line one')
    expect(firstItem.runs[breakIdx + 1].text).toBe('Line two')

    // 2. slide-builder layer: continuation run must use invisible bullet
    //    (breakAfter=true because this is not the last item)
    const textProps = toListTextProps(firstItem, false, true)

    // Group 0: 'Line one' — real bullet, ends with breakLine to close paragraph
    expect(textProps[0].text).toBe('Line one')
    expect(textProps[0].options?.bullet).toBe(true)
    expect(textProps[0].options?.indentLevel).toBe(0)
    expect(textProps[0].options?.breakLine).toBe(true)

    // Group 1: 'Line two' — invisible bullet gives correct marL
    expect(textProps[1].text).toBe('Line two')
    expect(textProps[1].options?.bullet).toEqual({ characterCode: '200B' })
    expect(textProps[1].options?.indentLevel).toBe(0)
    expect(textProps[1].options?.breakLine).toBe(true) // inter-item separator (breakAfter=true)

    restore()
  })

  it('ordered list with <br> also uses invisible bullet for continuation', () => {
    const { section } = setupSlide(`<ol id="ol"><li id="li">Step one<br>Detail</li></ol>`)
    const ol = section.querySelector('#ol')!
    const li = section.querySelector('#li')!

    mockRect(ol, { left: 50, top: 100, width: 600, height: 40 })
    mockRect(li, { left: 60, top: 100, width: 580, height: 40 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ol, { display: 'block', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px', backgroundColor: 'rgba(0,0,0,0)' }],
      [li, { display: 'list-item', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px', backgroundColor: 'rgba(0,0,0,0)' }],
    ])

    const slides = extractSlides()
    const listEl = slides[0].elements.find((e: any) => e.type === 'list') as any

    expect(listEl.ordered).toBe(true)

    const props = toListTextProps(listEl.items[0], true, false) // ordered, last item

    // First group: numbered bullet
    expect(props[0].options?.bullet).toEqual({ type: 'number', style: 'arabicPeriod' })
    expect(props[0].options?.breakLine).toBe(true)
    // Continuation: invisible bullet (not a number)
    expect(props[1].options?.bullet).toEqual({ characterCode: '200B' })
    expect(props[1].options?.breakLine).toBeUndefined()

    restore()
  })
})

// -----------------------------------------------------------------------
// T1: container borderStyle extraction (dashed / dotted)
// -----------------------------------------------------------------------

describe('container borderStyle extraction (via extractSlides)', () => {
  it('extracts dashed borderStyle from container div', () => {
    const { section } = setupSlide('<div id="box">Card</div>')
    const box = section.querySelector('#box')! as HTMLElement

    mockRect(box, { left: 50, top: 50, width: 400, height: 100 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [box, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left',
        backgroundColor: 'rgb(255,255,255)',
        borderTopWidth: '2px', borderTopStyle: 'dashed', borderTopColor: 'rgb(200,0,0)',
        borderRadius: '0px',
      }],
    ])

    const slides = extractSlides()
    restore()
    const container = slides[0].elements.find((e: any) => e.type === 'container') as any
    expect(container).toBeDefined()
    expect(container.style.borderStyle).toBe('dashed')
    expect(container.style.borderWidth).toBe(2)
    expect(container.style.borderColor).toBe('rgb(200,0,0)')
  })

  it('extracts dotted borderStyle from container div', () => {
    const { section } = setupSlide('<div id="box">Card</div>')
    const box = section.querySelector('#box')! as HTMLElement

    mockRect(box, { left: 50, top: 50, width: 400, height: 100 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [box, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left',
        backgroundColor: 'rgb(255,255,255)',
        borderTopWidth: '1px', borderTopStyle: 'dotted', borderTopColor: 'rgb(100,100,100)',
        borderRadius: '0px',
      }],
    ])

    const slides = extractSlides()
    restore()
    const container = slides[0].elements.find((e: any) => e.type === 'container') as any
    expect(container).toBeDefined()
    expect(container.style.borderStyle).toBe('dotted')
  })

  it('omits borderStyle for solid border (default)', () => {
    const { section } = setupSlide('<div id="box">Card</div>')
    const box = section.querySelector('#box')! as HTMLElement

    mockRect(box, { left: 50, top: 50, width: 400, height: 100 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [box, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left',
        backgroundColor: 'rgb(255,255,255)',
        borderTopWidth: '2px', borderTopStyle: 'solid', borderTopColor: 'rgb(0,0,0)',
        borderRadius: '0px',
      }],
    ])

    const slides = extractSlides()
    restore()
    const container = slides[0].elements.find((e: any) => e.type === 'container') as any
    expect(container).toBeDefined()
    // borderStyle is 'solid' but still propagated — slide-builder decides whether to apply dashType
    expect(container.style.borderWidth).toBe(2)
  })
})

// -----------------------------------------------------------------------
// T2: heading padding extraction
// -----------------------------------------------------------------------

describe('heading padding extraction (via extractSlides)', () => {
  it('extracts paddingLeft from heading with CSS padding', () => {
    const { section } = setupSlide('<h2 id="h">Decorated Heading</h2>')
    const h2 = section.querySelector('#h')! as HTMLElement

    mockRect(h2, { left: 30, top: 50, width: 600, height: 40 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [h2, {
        fontSize: '28px', fontWeight: '700', fontFamily: 'Arial',
        color: 'rgb(0,0,0)', lineHeight: '34px', textAlign: 'left',
        borderLeftWidth: '4px', borderLeftColor: 'rgb(39,174,96)',
        borderBottomWidth: '0px',
        paddingTop: '8px', paddingRight: '0px', paddingBottom: '8px', paddingLeft: '16px',
      }],
    ])

    const slides = extractSlides()
    restore()
    const heading = slides[0].elements.find((e: any) => e.type === 'heading') as any
    expect(heading).toBeDefined()
    expect(heading.style.paddingLeft).toBe(16)
    expect(heading.style.paddingTop).toBe(8)
  })

  it('omits padding properties when all zero', () => {
    const { section } = setupSlide('<h1 id="h">Plain Heading</h1>')
    const h1 = section.querySelector('#h')! as HTMLElement

    mockRect(h1, { left: 0, top: 0, width: 600, height: 50 })
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [h1, {
        fontSize: '40px', fontWeight: '700', fontFamily: 'Arial',
        color: 'rgb(0,0,0)', lineHeight: '48px', textAlign: 'left',
        borderLeftWidth: '0px', borderBottomWidth: '0px',
        paddingTop: '0px', paddingRight: '0px', paddingBottom: '0px', paddingLeft: '0px',
      }],
    ])

    const slides = extractSlides()
    restore()
    const heading = slides[0].elements.find((e: any) => e.type === 'heading') as any
    expect(heading).toBeDefined()
    expect(heading.style.paddingLeft).toBeUndefined()
    expect(heading.style.paddingTop).toBeUndefined()
  })
})

// -----------------------------------------------------------------------
// T5: display:inline span with borderRadius treated as badge
// -----------------------------------------------------------------------

describe('display:inline span with borderRadius as badge (via extractSlides)', () => {
  it('inline span with borderRadius > 6 and opaque background is emitted as container shape', () => {
    const { section } = setupSlide(`
      <p id="para"><span id="badge">Status</span> text</p>
    `)
    const para = section.querySelector('#para')! as HTMLElement
    const badge = section.querySelector('#badge')! as HTMLElement

    mockRect(para, { left: 50, top: 100, width: 600, height: 30 })
    mockRect(badge, { left: 50, top: 102, width: 60, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [para, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [badge, {
        display: 'inline', backgroundColor: 'rgb(76,175,80)',
        color: 'rgb(255,255,255)', borderRadius: '12px',
        fontSize: '12px', fontFamily: 'Arial', fontWeight: '700',
      }],
    ])

    const slides = extractSlides()
    restore()
    const containers = slides[0].elements.filter((e: any) => e.type === 'container')
    expect(containers.length).toBeGreaterThanOrEqual(1)
    expect((containers[0] as any).style.borderRadius).toBe(12)
    expect((containers[0] as any).style.backgroundColor).toBe('rgb(76,175,80)')
    // REGRESSION GUARD: must be a full container with text runs, NOT a bg-only shape.
    // A bg-only shape (runs === undefined) caused slide 34 badge misalignment.
    // This assertion would have caught that regression immediately.
    expect((containers[0] as any).runs).toBeDefined()
    expect((containers[0] as any).runs.length).toBeGreaterThan(0)
  })

  it('inline code element (borderRadius=6, semi-transparent bg) is NOT emitted as container', () => {
    // Marp default theme: <code> has border-radius ≈ 6px and rgba(129,139,152,0.12) bg
    // This must NOT be extracted as a badge shape (would create opaque grey block in PPTX)
    const { section } = setupSlide(`
      <p id="para">Background is <code id="code">some-value</code> text</p>
    `)
    const para = section.querySelector('#para')! as HTMLElement
    const code = section.querySelector('#code')! as HTMLElement

    mockRect(para, { left: 50, top: 100, width: 600, height: 30 })
    mockRect(code, { left: 190, top: 105, width: 120, height: 22 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [para, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [code, {
        display: 'inline', backgroundColor: 'rgba(129, 139, 152, 0.12)',
        color: 'rgb(0,0,0)', borderRadius: '6px',
        fontSize: '14px', fontFamily: 'monospace', fontWeight: '400',
      }],
    ])

    const slides = extractSlides()
    restore()
    // No container shapes should be emitted for inline code
    const containers = slides[0].elements.filter((e: any) => e.type === 'container')
    expect(containers).toHaveLength(0)
    // The code text must appear in the paragraph runs
    const paragraph = slides[0].elements.find((e: any) => e.type === 'paragraph') as any
    expect(paragraph).toBeDefined()
    const codeRun = paragraph.runs?.find((r: any) => r.text?.includes('some-value'))
    expect(codeRun).toBeDefined()
  })

  it('display:inline strong with borderRadius:4px (slide 56/58) stays as a text highlight instead of a badge shape', () => {
    // <p>Bold <strong style="background:#f1c40f;border-radius:4px">highlight</strong> here</p>
    // Semantic inline tags such as <strong> should remain run highlights.
    const { section } = setupSlide(`
      <p id="para">Bold <strong id="strong">highlight</strong> here</p>
    `)
    const para   = section.querySelector('#para')!   as HTMLElement
    const strong = section.querySelector('#strong')! as HTMLElement

    mockRect(para,   { left: 50, top: 100, width: 600, height: 30 })
    mockRect(strong, { left: 98, top: 102, width: 80, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [para, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [strong, {
        display: 'inline', backgroundColor: 'rgb(241,196,15)',
        color: 'rgb(0,0,0)', borderRadius: '4px',
        fontSize: '16px', fontFamily: 'Arial', fontWeight: '700',
      }],
    ])

    const slides = extractSlides()
    restore()

    const containers = slides[0].elements.filter((e: any) => e.type === 'container') as any[]
    expect(containers).toHaveLength(0)

    // Paragraph must contain "highlight" text WITH its original highlight.
    const paragraph = slides[0].elements.find((e: any) => e.type === 'paragraph') as any
    expect(paragraph).toBeDefined()
    const highlightRun = paragraph.runs?.find((r: any) => r.text?.includes('highlight'))
    expect(highlightRun).toBeDefined()
    expect(highlightRun.backgroundColor).toBe('rgb(241,196,15)')
  })

  it('display:inline-block code stays a text highlight instead of a detached badge shape', () => {
    const { section } = setupSlide(`
      <p id="para">Value: <code id="code">ABC-123</code> ready</p>
    `)
    const para = section.querySelector('#para')! as HTMLElement
    const code = section.querySelector('#code')! as HTMLElement

    mockRect(para, { left: 50, top: 100, width: 600, height: 30 })
    mockRect(code, { left: 118, top: 102, width: 92, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [para, {
        display: 'block', fontSize: '16px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [code, {
        display: 'inline-block', backgroundColor: 'rgb(241,196,15)',
        color: 'rgb(0,0,0)', borderRadius: '6px',
        fontSize: '14px', fontFamily: 'monospace', fontWeight: '400',
      }],
    ])

    const slides = extractSlides()
    restore()

    const containers = slides[0].elements.filter((e: any) => e.type === 'container')
    expect(containers).toHaveLength(0)

    const paragraph = slides[0].elements.find((e: any) => e.type === 'paragraph') as any
    expect(paragraph).toBeDefined()
    const codeRun = paragraph.runs?.find((r: any) => r.text?.includes('ABC-123'))
    expect(codeRun).toBeDefined()
    expect(codeRun.backgroundColor).toBe('rgb(241,196,15)')
  })
})

// -----------------------------------------------------------------------
// Pagination source detection
// We keep only the deck-wide source flag and let slide-builder.ts decide
// whether to enable PowerPoint's built-in slide-number field.
// -----------------------------------------------------------------------

describe('pagination source detection', () => {
  it('records sourceHasPagination when data-marpit-pagination is present and does not emit a page-number paragraph', () => {
    document.body.innerHTML = `
      <div id=":$p">
        <svg data-marpit-svg="" viewBox="0 0 1280 720">
          <foreignObject width="1280" height="720">
            <section id="1" data-marpit-pagination="5">
              <h1 id="h1">Title</h1>
            </section>
          </foreignObject>
        </svg>
      </div>
    `
    const section = document.querySelector('section')!
    const h1 = section.querySelector('h1')! as HTMLElement

    mockRect(section, { left: 0, top: 0, width: 1280, height: 720 })
    mockRect(h1, { left: 70, top: 80, width: 600, height: 50 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [h1, { fontSize: '40px', fontWeight: '700', color: 'rgb(0,0,0)', fontFamily: 'Arial' }],
    ], [
      [section, '::after', {
        content: '"5"',
        display: 'block',
        position: 'absolute',
        right: '0px',
        bottom: '0px',
        paddingRight: '40px',
        paddingBottom: '30px',
        color: 'rgb(119, 119, 119)',
        fontSize: '18px',
        lineHeight: '24px',
        textAlign: 'right',
        opacity: '1',
        visibility: 'visible',
      }],
    ])

    const slides = extractSlides()
    restore()

    expect(slides[0].sourceHasPagination).toBe(true)
    const numberParagraphs = slides[0].elements.filter(
      (e: any) => e.type === 'paragraph' && e.runs?.some((r: any) => r.text === '5')
    )
    expect(numberParagraphs).toHaveLength(0)
  })

  it('preserves pagination pseudo backgrounds as shapes while leaving numbering to the native slide number field', () => {
    document.body.innerHTML = `
      <div id=":$p">
        <svg data-marpit-svg="" viewBox="0 0 1280 720">
          <foreignObject width="1280" height="720">
            <section id="1" data-marpit-pagination="12">
              <h1 id="h1">Title</h1>
            </section>
          </foreignObject>
        </svg>
      </div>
    `
    const section = document.querySelector('section')!
    const h1 = section.querySelector('h1')! as HTMLElement

    mockRect(section, { left: 0, top: 0, width: 1280, height: 720 })
    mockRect(h1, { left: 70, top: 80, width: 600, height: 50 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [h1, { fontSize: '40px', fontWeight: '700', color: 'rgb(0,0,0)', fontFamily: 'Arial' }],
    ], [
      [section, '::after', {
        content: '"12"',
        display: 'block',
        position: 'absolute',
        right: '0px',
        bottom: '0px',
        width: '96px',
        height: '28px',
        backgroundColor: 'rgb(34,68,102)',
        color: 'rgb(255,255,255)',
      }],
    ])

    const slides = extractSlides()
    restore()

    expect(slides[0].sourceHasPagination).toBe(true)
    const containers = slides[0].elements.filter((e: any) => e.type === 'container')
    expect(containers).toHaveLength(1)
    expect(containers[0]).toMatchObject({
      x: 0,
      y: 692,
      width: 96,
      height: 28,
      style: { backgroundColor: 'rgb(34,68,102)' },
    })
  })

  it('records sourceHasPagination even when the pseudo-element is hidden', () => {
    document.body.innerHTML = `
      <div id=":$p">
        <svg data-marpit-svg="" viewBox="0 0 1280 720">
          <foreignObject width="1280" height="720">
            <section id="52" data-marpit-pagination="52">
              <p id="p">Content</p>
            </section>
          </foreignObject>
        </svg>
      </div>
    `

    const section = document.querySelector('section')!
    const p = section.querySelector('#p')! as HTMLElement
    mockRect(section, { left: 0, top: 0, width: 1280, height: 720 })
    mockRect(p, { left: 50, top: 100, width: 600, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {
        display: 'block',
        fontSize: '16px',
        fontFamily: 'Arial',
        color: 'rgb(0,0,0)',
        fontWeight: '400',
        lineHeight: '24px',
        textAlign: 'left',
      }],
    ], [
      [section, '::after', {
        content: '"52"',
        display: 'none',
        position: 'absolute',
        right: '0px',
        bottom: '0px',
        color: 'rgb(119, 119, 119)',
        fontSize: '18px',
      }],
    ])

    const slides = extractSlides()
    restore()

    expect(slides[0].sourceHasPagination).toBe(true)
  })

  it('records sourceHasPagination as false when data-marpit-pagination is absent', () => {
    const { section } = setupSlide('<p id="p">Content</p>')
    const p = section.querySelector('p')! as HTMLElement
    mockRect(p, { left: 50, top: 100, width: 600, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p, {
        display: 'block',
        fontSize: '16px',
        fontFamily: 'Arial',
        color: 'rgb(0,0,0)',
        fontWeight: '400',
        lineHeight: '24px',
        textAlign: 'left',
      }],
    ])

    const slides = extractSlides()
    restore()

    expect(slides[0].sourceHasPagination).toBe(false)
  })
})

describe('list badge extraction', () => {
  it('records leadingOffset for list items when a leading badge span is extracted as a shape', () => {
    const { section } = setupSlide(`
      <ul id="list"><li id="item"><span id="badge">NEW</span> Launch</li></ul>
    `)
    const list = section.querySelector('#list')! as HTMLElement
    const item = section.querySelector('#item')! as HTMLElement
    const badge = section.querySelector('#badge')! as HTMLElement

    mockRect(list, { left: 70, top: 150, width: 500, height: 40 })
    mockRect(item, { left: 100, top: 150, width: 430, height: 28 })
    mockRect(badge, { left: 100, top: 152, width: 64, height: 24 })

    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [list, {
        display: 'block', fontSize: '18px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '27px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [item, {
        display: 'list-item', fontSize: '18px', fontFamily: 'Arial',
        fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '27px',
        textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      }],
      [badge, {
        display: 'inline', backgroundColor: 'rgb(76,175,80)',
        color: 'rgb(255,255,255)', borderRadius: '12px',
        fontSize: '12px', fontFamily: 'Arial', fontWeight: '700',
      }],
    ])

    const slides = extractSlides()
    restore()

    const containers = slides[0].elements.filter((e: any) => e.type === 'container') as any[]
    expect(containers.length).toBeGreaterThanOrEqual(1)

    const listEl = slides[0].elements.find((e: any) => e.type === 'list') as any
    expect(listEl).toBeDefined()
    expect(listEl.items).toHaveLength(1)
    expect(listEl.items[0].leadingOffset).toBeCloseTo(64, 4)
  })
})

// -----------------------------------------------------------------------
// inline badge inside <li> — slide 36 regression
// <li>Review <span style="border-radius:8px;background:#c05621">Needs review</span></li>
// The badge span must be extracted as a container shape (rounded corners),
// NOT as a flat inline highlight in the list item runs.
// -----------------------------------------------------------------------

describe('inline badge inside <li> extracted as container shape (slide 36)', () => {
  it('display:inline span with border-radius>6 and opaque bg inside <li> becomes a container shape', () => {
    const { section } = setupSlide(`
      <ol id="ol">
        <li id="li1">Design item</li>
        <li id="li2">Review <span id="badge">Needs review</span></li>
      </ol>
    `)
    const ol   = section.querySelector('#ol')!   as HTMLElement
    const li1  = section.querySelector('#li1')!  as HTMLElement
    const li2  = section.querySelector('#li2')!  as HTMLElement
    const badge = section.querySelector('#badge')! as HTMLElement

    mockRect(ol,    { left: 60, top: 200, width: 800, height: 80 })
    mockRect(li1,   { left: 60, top: 200, width: 800, height: 36 })
    mockRect(li2,   { left: 60, top: 236, width: 800, height: 36 })
    mockRect(badge, { left: 210, top: 242, width: 90, height: 22 })

    const liStyle = {
      display: 'list-item', fontSize: '16px', fontFamily: 'Arial',
      fontWeight: '400', color: 'rgb(0,0,0)', lineHeight: '24px',
      textAlign: 'left', backgroundColor: 'rgba(0,0,0,0)',
      fontStyle: 'normal',
    }
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [ol, { display: 'block', fontSize: '16px', fontFamily: 'Arial', color: 'rgb(0,0,0)', fontWeight: '400', fontStyle: 'normal', textAlign: 'left', lineHeight: '24px' }],
      [li1, liStyle],
      [li2, liStyle],
      [badge, {
        display: 'inline',
        backgroundColor: 'rgb(192,86,33)',
        color: 'rgb(255,255,255)',
        borderRadius: '8px',
        fontSize: '12.8px', fontFamily: 'Arial', fontWeight: '400',
        fontStyle: 'normal',
      }],
    ])

    const slides = extractSlides()
    restore()
    const els = slides[0].elements

    // Container shape must be emitted (rounded badge)
    const containers = els.filter((e: any) => e.type === 'container')
    expect(containers).toHaveLength(1)
    expect((containers[0] as any).style.backgroundColor).toBe('rgb(192,86,33)')
    expect((containers[0] as any).style.borderRadius).toBe(8)

    // Badge runs must NOT appear with backgroundColor in the list item runs
    // (text is rendered inside the shape, not as inline highlight)
    const listEl = els.find((e: any) => e.type === 'list') as any
    expect(listEl).toBeDefined()
    const item2 = listEl.items.find((i: any) => i.text?.includes('Review'))
    expect(item2).toBeDefined()
    const badgeRunWithHighlight = item2.runs?.find(
      (r: any) => r.backgroundColor === 'rgb(192,86,33)'
    )
    // Badge text should NOT be in runs as a flat highlight — it's in the shape
    expect(badgeRunWithHighlight).toBeUndefined()
  })
})

// -----------------------------------------------------------------------
// non-leading inline-flex badge in <p> — slide 34 regression
// <p><span "1" (leading)> Install <span "2" (non-leading)> Create config</p>
// Badge "1" (leading) → container shape WITH runs, text skipped from paragraph
// Badge "2" (non-leading) → container shape WITHOUT runs (bg-only), text kept
//   in paragraph run with correct text color but NO backgroundColor (flat box)
// -----------------------------------------------------------------------

describe('non-leading inline-flex badge in <p> extracted as bg-only shape (slide 34)', () => {
  it('leading badge gets runs; non-leading badge is bg-only; paragraph run has text without backgroundColor', () => {
    const { section } = setupSlide(`
      <p id="p1">
        <span id="badge1">1</span> Install
        <span id="badge2">2</span> Create config
      </p>
    `)
    const p1     = section.querySelector('#p1')!     as HTMLElement
    const badge1 = section.querySelector('#badge1')! as HTMLElement
    const badge2 = section.querySelector('#badge2')! as HTMLElement

    // badge1 is flush at the paragraph left (leading position)
    // badge2 is 160 px to the right (non-leading)
    mockRect(p1,     { left: 60, top: 200, width: 700, height: 30 })
    mockRect(badge1, { left: 60, top: 203, width: 28, height: 28 })
    mockRect(badge2, { left: 220, top: 203, width: 28, height: 28 })

    const badgeStyle = {
      display: 'inline-flex',
      backgroundColor: 'rgb(49,130,206)',
      color: 'rgb(255,255,255)',
      borderRadius: '14px',
      fontSize: '16px', fontFamily: 'Arial', fontWeight: '700',
      fontStyle: 'normal',
    }
    const restore = mockStyles([
      [section, { backgroundColor: 'rgb(255,255,255)' }],
      [p1, {
        display: 'block',
        fontSize: '18px', fontFamily: 'Arial', fontWeight: '400',
        color: 'rgb(30,41,59)', fontStyle: 'normal',
        textAlign: 'left', lineHeight: '27px',
        backgroundColor: 'rgba(0,0,0,0)',
      }],
      [badge1, badgeStyle],
      [badge2, badgeStyle],
    ])

    const slides = extractSlides()
    restore()
    const els = slides[0].elements

    const containers = els.filter((e: any) => e.type === 'container')
    // Both badges should produce container shapes
    expect(containers).toHaveLength(2)

    // Leading badge (badge1) has runs with text "1"
    const leadingShape = containers.find(
      (c: any) => Math.abs(c.x - (60 - 0)) < 5
    ) as any
    expect(leadingShape).toBeDefined()
    expect(leadingShape.runs).toBeDefined()
    expect(leadingShape.runs?.some((r: any) => r.text === '1')).toBe(true)

    // Non-leading badge (badge2) is a background-only shape (no runs / empty runs)
    const bgOnlyShape = containers.find(
      (c: any) => Math.abs(c.x - (220 - 0)) < 5
    ) as any
    expect(bgOnlyShape).toBeDefined()
    // bg-only shape has no text runs (or runs is undefined)
    const bgOnlyHasText = bgOnlyShape.runs?.some(
      (r: any) => !r.breakLine && r.text?.trim() !== ''
    )
    expect(bgOnlyHasText).toBeFalsy()

    // Paragraph element exists
    const para = els.find((e: any) => e.type === 'paragraph') as any
    expect(para).toBeDefined()

    // "2" text is present in paragraph runs for correct text flow
    const run2 = para?.runs?.find((r: any) => r.text?.trim() === '2')
    expect(run2).toBeDefined()
    // "2" run must NOT have a backgroundColor (bg-only shape handles the visual)
    expect(run2?.backgroundColor).toBeUndefined()

    // "1" text must NOT be in paragraph runs (it's rendered inside the leading shape)
    const run1 = para?.runs?.find((r: any) => r.text?.trim() === '1')
    expect(run1).toBeUndefined()
  })
})
