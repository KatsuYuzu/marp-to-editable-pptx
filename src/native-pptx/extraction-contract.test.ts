import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import type { SlideData, SlideElement } from './types'
import { findMissingContractText, normalizeContractText } from './text-contract'

const PROJECT_ROOT = path.resolve(__dirname, '../..')
const HTML_PATH =
  process.env.MARP_STRUCTURAL_HTML ??
  path.join(
    PROJECT_ROOT,
    'src',
    'native-pptx',
    'test-fixtures',
    'slides-ci.html',
  )
const DATA_PATH =
  process.env.MARP_STRUCTURAL_DATA ??
  path.join(
    PROJECT_ROOT,
    'src',
    'native-pptx',
    'test-fixtures',
    'slides-ci.json',
  )
const BROWSER_PATH =
  process.env.CHROME_PATH ??
  [
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
    process.env.LOCALAPPDATA
      ? path.join(
          process.env.LOCALAPPDATA,
          'Microsoft',
          'Edge',
          'Application',
          'msedge.exe',
        )
      : undefined,
  ].find((candidate): candidate is string => !!candidate && fs.existsSync(candidate))

const CAN_RUN = !!(
  BROWSER_PATH &&
  fs.existsSync(HTML_PATH) &&
  fs.existsSync(DATA_PATH)
)
const describeOrSkip = CAN_RUN ? describe : describe.skip

function collectElementText(element: SlideElement): string[] {
  const texts: string[] = []
  if ('runs' in element && Array.isArray(element.runs)) {
    texts.push(
      ...element.runs
        .filter((run) => !run.breakLine)
        .map((run) => run.text),
    )
  }
  if (element.type === 'code' && element.runs.length === 0) {
    texts.push(element.text)
  }
  if (element.type === 'list') {
    for (const item of element.items) {
      texts.push(
        ...(item.runs ?? [])
          .filter((run) => !run.breakLine)
          .map((run) => run.text),
      )
    }
  }
  if (element.type === 'table') {
    for (const row of element.rows) {
      for (const cell of row.cells) {
        const paragraphs = cell.paragraphs?.flatMap((paragraph) => paragraph.runs)
        const runs = paragraphs ?? cell.runs
        texts.push(
          ...runs.filter((run) => !run.breakLine).map((run) => run.text),
        )
      }
    }
  }
  if (element.type === 'container') {
    for (const child of element.children) texts.push(...collectElementText(child))
  }
  return texts.map(normalizeContractText).filter(Boolean)
}

describeOrSkip('all fixture slides preserve browser-rendered text', () => {
  let slides: SlideData[]
  let expectedTexts: string[][]

  beforeAll(async () => {
    slides = JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) as SlideData[]

    const puppeteer = require('puppeteer-core')
    const browser = await puppeteer.launch({
      executablePath: BROWSER_PATH,
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu'],
    })
    try {
      const page = await browser.newPage()
      await page.setViewport({ width: 1280, height: 720 })
      await page.goto(pathToFileURL(HTML_PATH).href, { waitUntil: 'networkidle0' })
      await new Promise((resolve) => setTimeout(resolve, 500))

      expectedTexts = await page.evaluate(() => {
        function visibleTextNodes(section: Element): string[] {
          const walker = document.createTreeWalker(
            section,
            NodeFilter.SHOW_TEXT,
          )
          const texts: string[] = []
          let node: Node | null
          while ((node = walker.nextNode())) {
            const text = node.textContent?.trim() ?? ''
            if (!text) continue

            const parent = node.parentElement
            if (!parent) continue
            if (
              parent.closest(
                'script,style,[data-marpit-presenter-notes],.bespoke-marp-note',
              ) ||
              parent.closest('.katex-mathml') ||
              parent.closest('p')?.querySelector('mjx-container')
            )
            continue
            let nestedAncestor: Element | null = parent
            let insideNestedSvg = false
            while (nestedAncestor && nestedAncestor.tagName.toLowerCase() !== 'section') {
              if (nestedAncestor.tagName.toLowerCase() === 'svg') {
                insideNestedSvg = true
                break
              }
              nestedAncestor = nestedAncestor.parentElement
            }
            if (insideNestedSvg) continue
            const style = getComputedStyle(parent)
            if (style.display === 'none' || style.visibility === 'hidden')
              continue

            const range = document.createRange()
            range.selectNodeContents(node)
            const rect = range.getBoundingClientRect()
            if (rect.width === 0 || rect.height === 0) continue

            texts.push(text.replace(/\s+/gu, ''))
          }
          return texts
        }

        return Array.from(document.querySelectorAll('svg[data-marpit-svg]')).map(
          (svg) => {
            const sections = Array.from(svg.querySelectorAll(':scope > foreignObject > section'))
            const content =
              sections.find(
                (section) =>
                  section.getAttribute('data-marpit-advanced-background') ===
                  'content',
              ) ?? sections[0]
            return content ? visibleTextNodes(content) : []
          },
        )
      })
    } finally {
      await browser.close()
    }
  }, 120000)

  it('keeps all 96 fixture slides and their order', () => {
    expect(expectedTexts).toHaveLength(96)
    expect(slides).toHaveLength(96)
  })

  it('keeps every visible browser text node in extraction output', () => {
    const failures: string[] = []
    for (let slideIndex = 0; slideIndex < expectedTexts.length; slideIndex++) {
      const actual = normalizeContractText(
        slides[slideIndex].elements.flatMap(collectElementText).join(''),
      )
      const missing = findMissingContractText(expectedTexts[slideIndex], actual)
      if (missing) {
        failures.push(
          `slide ${slideIndex + 1}: missing or reordered text "${missing.slice(0, 80)}"`,
        )
      }
    }
    expect(failures).toEqual([])
  })
})
