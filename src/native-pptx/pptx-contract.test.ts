import fs from 'node:fs'
import path from 'node:path'
import JSZip from 'jszip'
import type { SlideData, SlideElement } from './types'
import { createPptxRenderPlan } from './pptx-plan'

const PROJECT_ROOT = path.resolve(__dirname, '../..')
const PPTX_PATH =
  process.env.MARP_PPTX_PATH ??
  path.join(PROJECT_ROOT, 'dist', 'slides-ci.pptx')
const DATA_PATH =
  process.env.MARP_STRUCTURAL_DATA ??
  path.join(
    PROJECT_ROOT,
    'src',
    'native-pptx',
    'test-fixtures',
    'slides-ci.json',
  )
const CAN_RUN = fs.existsSync(PPTX_PATH) && fs.existsSync(DATA_PATH)
const describeOrSkip = CAN_RUN ? describe : describe.skip

function normalizeText(text: string): string {
  return text.replace(/\s+|\ufe0f/gu, '')
}

function collectSlideText(elements: SlideElement[]): string[] {
  const chunks: string[] = []
  for (const element of elements) {
    if ('runs' in element && Array.isArray(element.runs)) {
      chunks.push(
        ...element.runs
          .filter((run) => !run.breakLine)
          .map((run) => run.text),
      )
    }
    if (element.type === 'list') {
      for (const item of element.items) {
        chunks.push(
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
          chunks.push(
            ...runs.filter((run) => !run.breakLine).map((run) => run.text),
          )
        }
      }
    }
    if (element.type === 'container') {
      chunks.push(...collectSlideText(element.children))
    }
  }
  return chunks
    .flatMap((text) => text.match(/[\p{L}\p{N}-]+/gu) ?? [])
    .map(normalizeText)
    .filter((text) => text.length >= 2)
}

describeOrSkip('SlideData to PPTX XML contract', () => {
  let expectedSlides: SlideData[]
  let pptxText: string[]

  beforeAll(async () => {
    expectedSlides = JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'))
    const zip = await JSZip.loadAsync(fs.readFileSync(PPTX_PATH))
    const slideEntries = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort(
        (left, right) =>
          Number(left.match(/\d+/)?.[0] ?? 0) -
          Number(right.match(/\d+/)?.[0] ?? 0),
      )

    pptxText = await Promise.all(
      slideEntries.map(async (name) => {
        const xml = await zip.file(name)!.async('string')
        return normalizeText(
          Array.from(xml.matchAll(/<a:t(?: [^>]*)?>([\s\S]*?)<\/a:t>/g))
            .map((match) =>
              match[1]
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&apos;/g, "'")
                .replace(/&amp;/g, '&')
                .replace(/&#([0-9]+);/g, (_, decimal) =>
                  String.fromCodePoint(Number.parseInt(decimal, 10)),
                )
                .replace(/&#x([0-9a-f]+);/gi, (_, hex) =>
                  String.fromCodePoint(Number.parseInt(hex, 16)),
                ),
            )
            .join(''),
        )
      }),
    )
  })

  it('keeps the same slide count', () => {
    expect(pptxText).toHaveLength(expectedSlides.length)
  })

  it('keeps each extracted text run in its slide XML', () => {
    const failures: string[] = []
    const plan = createPptxRenderPlan(expectedSlides)
    plan.slides.forEach((slidePlan, index) => {
      const expected = collectSlideText(slidePlan.visibleElements)
      const actual = pptxText[index]
      const missing = expected.find((text) => !actual.includes(text))
      if (missing) {
        failures.push(
          `slide ${index + 1}: missing "${missing.slice(0, 80)}" in PPTX XML`,
        )
      }
    })
    expect(failures).toEqual([])
  })
})
