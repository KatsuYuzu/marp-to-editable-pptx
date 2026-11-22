import fs from 'node:fs'
import path from 'node:path'
import JSZip from 'jszip'
import type { SlideData, SlideElement } from './types'
import { createPptxRenderPlan } from './pptx-plan'
import { findMissingContractText, normalizeContractText } from './text-contract'

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
const MANIFEST_PATH = path.join(
  PROJECT_ROOT,
  'src',
  'native-pptx',
  'test-fixtures',
  'fixture-manifest.json',
)

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
    .map(normalizeContractText)
    .filter(Boolean)
}

function collectElementsWithVisibleSourceText(
  elements: SlideElement[],
  slideHeight: number,
): SlideElement[] {
  const visible: SlideElement[] = []
  for (const element of elements) {
    if (
      element.type !== 'container' &&
      element.type !== 'image' &&
      slideHeight - element.y < 20
    ) {
      continue
    }
    if (element.type === 'container') {
      visible.push({
        ...element,
        children: collectElementsWithVisibleSourceText(
          element.children,
          slideHeight,
        ),
      })
    } else {
      visible.push(element)
    }
  }
  return visible
}

function collectElements(elements: SlideElement[]): SlideElement[] {
  return elements.flatMap((element) => [
    element,
    ...(element.type === 'container'
      ? collectElements(element.children)
      : []),
  ])
}

function collectVisibleSourceText(
  elements: SlideElement[],
  slideHeight: number,
): string[] {
  return collectSlideText(
    collectElementsWithVisibleSourceText(elements, slideHeight),
  )
}

describeOrSkip('SlideData to PPTX XML contract', () => {
  let expectedSlides: SlideData[]
  let pptxText: string[]
  let pptxShapeCounts: number[]
  let pptxXml: string[]

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

    const extracted = await Promise.all(
      slideEntries.map(async (name) => {
        const xml = await zip.file(name)!.async('string')
        return {
          text: normalizeContractText(
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
          ),
          shapeCount: (xml.match(/<p:sp(?:>| )/g) ?? []).length,
        }
      }),
    )
    pptxText = extracted.map((entry) => entry.text)
    pptxShapeCounts = extracted.map((entry) => entry.shapeCount)
    pptxXml = await Promise.all(
      slideEntries.map(async (name) => zip.file(name)!.async('string')),
    )
  })

  it('keeps the same slide count', () => {
    expect(pptxText).toHaveLength(expectedSlides.length)
  })

  it('keeps each extracted text run in its slide XML', () => {
    const failures: string[] = []
    expectedSlides.forEach((slide, index) => {
      const expected = collectVisibleSourceText(slide.elements, slide.height)
      const actual = pptxText[index]
      const missing = findMissingContractText(expected, actual)
      if (missing) {
        failures.push(
          `slide ${index + 1}: missing "${missing.slice(0, 80)}" in PPTX XML`,
        )
      }
    })
    expect(failures).toEqual([])
  })

  it('emits PPTX shapes whenever the plan has editable content', () => {
    const plan = createPptxRenderPlan(expectedSlides)
    plan.slides.forEach((slide, index) => {
      if (slide.editableElementCount > 0) {
        expect(pptxShapeCounts[index]).toBeGreaterThan(0)
      }
    })
  })

  it('preserves native editable list and table semantics by fixture category', () => {
    const manifest = JSON.parse(fs.readFileSync(MANIFEST_PATH, 'utf8'))
    const failures: string[] = []
    manifest.slides.forEach((entry: any, index: number) => {
      const xml = pptxXml[index]
      const expectedTypes = collectElements(
        expectedSlides[index].elements,
      ).map((element) => element.type)
      if (
        expectedTypes.includes('table') &&
        !xml.includes('<a:tbl>')
      ) {
        failures.push(`${entry.id}: expected native table`)
      }
      if (
        expectedTypes.includes('list') &&
        !xml.includes('<a:buChar') &&
        !xml.includes('<a:buAutoNum')
      ) {
        failures.push(`${entry.id}: expected native list marker`)
      }
      if (
        expectedTypes.includes('code') &&
        (xml.match(/<p:sp(?:>| )/g) ?? []).length === 0
      ) {
        failures.push(`${entry.id}: expected editable code shape`)
      }
    })
    expect(failures).toEqual([])
  })

  it('preserves hyperlinks as clickable PPTX relationships', () => {
    const expectedHyperlinks = expectedSlides.map((slide) => {
      const runs: any[] = []
      const collect = (elements: SlideElement[]) => {
        for (const element of elements) {
          if ('runs' in element && Array.isArray(element.runs))
            runs.push(...element.runs)
          if (element.type === 'list')
            for (const item of element.items) runs.push(...item.runs)
          if (element.type === 'table')
            for (const row of element.rows)
              for (const cell of row.cells)
                runs.push(
                  ...(cell.paragraphs?.flatMap(
                    (paragraph) => paragraph.runs,
                  ) ?? cell.runs),
                )
          if (element.type === 'container') collect(element.children)
        }
      }
      collect(slide.elements)
      return runs.filter((run) => !!run.hyperlink).length
    })
    pptxXml.forEach((xml, index) => {
      const actual = (xml.match(/<a:hlinkClick\b/g) ?? []).length
      expect(actual).toBe(expectedHyperlinks[index])
    })
  })
})
