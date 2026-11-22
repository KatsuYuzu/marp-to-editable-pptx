import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import type { SlideData, SlideElement } from './types'
import { findMissingContractText, normalizeContractText } from './text-contract'

const root = path.resolve(__dirname, '../..')
const bespokePath =
  process.env.MARP_BESPOKE_HTML ??
  path.join(root, 'src', 'native-pptx', 'test-fixtures', 'slides-bespoke-ci.html')
const dataPath =
  process.env.MARP_STRUCTURAL_DATA ??
  path.join(root, 'src', 'native-pptx', 'test-fixtures', 'slides-ci.json')
const browserPath =
  process.env.CHROME_PATH ??
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
const canRun =
  fs.existsSync(bespokePath) &&
  fs.existsSync(dataPath) &&
  fs.existsSync(browserPath)
const describeOrSkip = canRun ? describe : describe.skip

function collectText(elements: SlideElement[]): string[] {
  const result: string[] = []
  for (const element of elements) {
    if ('runs' in element && Array.isArray(element.runs)) {
      result.push(
        ...element.runs
          .filter((run) => !run.breakLine)
          .map((run) => run.text),
      )
    }
    if (element.type === 'list') {
      for (const item of element.items) {
        result.push(
          ...(item.runs ?? [])
            .filter((run) => !run.breakLine)
            .map((run) => run.text),
        )
      }
    }
    if (element.type === 'table') {
      for (const row of element.rows) {
        for (const cell of row.cells) {
          const runs =
            cell.paragraphs?.flatMap((paragraph) => paragraph.runs) ?? cell.runs
          result.push(
            ...runs.filter((run) => !run.breakLine).map((run) => run.text),
          )
        }
      }
    }
    if (element.type === 'container') result.push(...collectText(element.children))
  }
  return result.map(normalizeContractText).filter(Boolean)
}

describeOrSkip('bespoke active-slide contract', () => {
  let actual: Array<{ text: string[]; active: string; transforms: number }>

  beforeAll(async () => {
    const puppeteer = require('puppeteer-core')
    const browser = await puppeteer.launch({
      executablePath: browserPath,
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    })
    try {
      const page = await browser.newPage()
      await page.setViewport({ width: 1280, height: 720 })
      await page.goto(pathToFileURL(bespokePath).href, {
        waitUntil: 'networkidle0',
      })
      actual = []
      for (let pageIndex = 1; pageIndex <= 97; pageIndex++) {
        for (let attempt = 0; attempt < 8 && pageIndex > 1; attempt++) {
          const activeId = await page.evaluate(() => {
            const section = Array.from(document.querySelectorAll('section')).find(
              (candidate) =>
                getComputedStyle(candidate).pointerEvents === 'auto' &&
                candidate.getBoundingClientRect().width > 0,
            )
            return section
              ? Array.from(
                  document.querySelectorAll('svg[data-marpit-svg]'),
                ).indexOf(section.closest('svg') as SVGSVGElement) + 1
              : 0
          })
          if (activeId === pageIndex) break
          if (activeId > pageIndex) {
            throw new Error(
              `Bespoke navigation skipped page ${pageIndex} (active ${activeId})`,
            )
          }
          await page.keyboard.press('ArrowRight')
        }
        await page.waitForFunction((expected: number) => {
          const section = Array.from(document.querySelectorAll('section')).find(
            (candidate) =>
              getComputedStyle(candidate).pointerEvents === 'auto' &&
              candidate.getBoundingClientRect().width > 0,
          )
          return (
            Array.from(
              document.querySelectorAll('svg[data-marpit-svg]'),
            ).indexOf(section?.closest('svg') as SVGSVGElement) +
              1 ===
            expected
          )
        }, undefined, pageIndex)
        actual.push(
          await page.evaluate(() => {
            const activeLayer = Array.from(
              document.querySelectorAll('section'),
            ).find(
              (section) =>
                getComputedStyle(section).pointerEvents === 'auto' &&
                section.getBoundingClientRect().width > 0,
            )!
            const svg = activeLayer.closest('svg')!
            const active =
              Array.from(
                svg.querySelectorAll<HTMLElement>(
                  ':scope > foreignObject > section',
                ),
              ).find(
                (section) =>
                  section.getAttribute(
                    'data-marpit-advanced-background',
                  ) === 'content',
              ) ??
              (svg.querySelector('section') as HTMLElement)
            const mainSvg = svg
            const text: string[] = []
            const walker = document.createTreeWalker(active, NodeFilter.SHOW_TEXT)
            let node: Node | null
            while ((node = walker.nextNode())) {
              const parent = node.parentElement
              const value = node.textContent?.trim() ?? ''
              if (!parent || !value) continue
              if (
                parent.closest(
                  'script,style,[data-marpit-presenter-notes],.bespoke-marp-note,.katex-mathml',
                )
              )
                continue
              if (parent.closest('.katex')) continue
              if (
                parent.closest('p,li')?.querySelector('mjx-container')
              )
                continue
              if (parent.closest('svg') !== mainSvg) continue
              const style = getComputedStyle(parent)
              if (style.display === 'none' || style.visibility === 'hidden')
                continue
              const range = document.createRange()
              range.selectNodeContents(node)
              const rect = range.getBoundingClientRect()
              if (rect.width <= 0 || rect.height <= 0) continue
              text.push(value.replace(/\s+|\ufe0f/gu, ''))
            }
            return {
              active: active.getAttribute('data-marpit-pagination') ?? active.id,
              text,
              transforms: Array.from(
                active.querySelectorAll('*'),
              ).filter(
                (element) => getComputedStyle(element).transform !== 'none',
              ).length,
            }
          }),
        )
      }
    } finally {
      await browser.close()
    }
  }, 120000)

  it('preserves every static fixture page in active navigation', () => {
    const slides = JSON.parse(fs.readFileSync(dataPath, 'utf8')) as SlideData[]
    expect(actual.map((entry) => entry.active)).toEqual(
      Array.from({ length: 97 }, (_, index) => String(index + 1)),
    )
    const failures: string[] = []
    actual.forEach((entry, index) => {
      const missing = findMissingContractText(
        entry.text,
        collectText(slides[index].elements).join(''),
      )
      if (missing) failures.push(`page ${index + 1}: ${missing}`)
    })
    expect(failures).toEqual([])
  })
})
