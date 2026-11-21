import fs from 'node:fs'
import path from 'node:path'
import type { SlideData, SlideElement } from './types'

const root = path.resolve(__dirname, '../..')
const dataPath =
  process.env.MARP_STRUCTURAL_DATA ??
  path.join(root, 'src', 'native-pptx', 'test-fixtures', 'slides-ci.json')
const fixturePath = path.join(
  root,
  'src',
  'native-pptx',
  'test-fixtures',
  'pptx-export.md',
)
const manifestPath = path.join(
  root,
  'src',
  'native-pptx',
  'test-fixtures',
  'fixture-manifest.json',
)
const canRead = fs.existsSync(dataPath) && fs.existsSync(manifestPath)
const describeOrSkip = canRead ? describe : describe.skip

function flatten(elements: SlideElement[]): SlideElement[] {
  return elements.flatMap((element) => [
    element,
    ...(element.type === 'container' ? flatten(element.children) : []),
  ])
}

describeOrSkip('96-slide fixture integrity', () => {
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
  const slides = JSON.parse(fs.readFileSync(dataPath, 'utf8')) as SlideData[]

  it('has stable unique IDs for all 96 slides', () => {
    expect(manifest.slides).toHaveLength(96)
    expect(
      new Set(manifest.slides.map((slide: any) => slide.id)).size,
    ).toBe(96)
    expect(manifest.slides.map((slide: any) => slide.page)).toEqual(
      Array.from({ length: 96 }, (_, index) => index + 1),
    )
  })

  it('keeps visible text and shape geometry inside the canvas', () => {
    const failures: string[] = []
    slides.forEach((slide, index) => {
      for (const element of flatten(slide.elements)) {
        if (element.type === 'image' && element.rasterize) continue
        if (
          element.x < -0.5 ||
          element.y < -0.5 ||
          element.x + element.width > slide.width + 0.5 ||
          element.y + element.height > slide.height + 0.5
        ) {
          failures.push(
            `${index + 1}/${manifest.slides[index]?.id}: ${element.type} ` +
              `at ${element.x.toFixed(1)},${element.y.toFixed(1)} size ` +
              `${element.width.toFixed(1)}x${element.height.toFixed(1)}`,
          )
        }
      }
    })
    expect(failures).toEqual([])
  })

  it('uses only existing local assets and no remote fixture resources', () => {
    const markdown = fs.readFileSync(fixturePath, 'utf8')
    const references = Array.from(
      markdown.matchAll(/!\[[^\]]*\]\(([^)]+)\)|url\(['"]?([^)'"]+)/g),
      (match) => match[1] ?? match[2],
    )
    references.push(
      ...Array.from(
        markdown.matchAll(/<img\b[^>]*\bsrc=["']([^"']+)["'][^>]*>/g),
        (match) => match[1],
      ),
      ...Array.from(
        markdown.matchAll(
          /<(?:script|link)\b[^>]*(?:src|href)=["']([^"']+)["'][^>]*>/g,
        ),
        (match) => match[1],
      ),
    )
    const remote = references.filter((reference) =>
      /^https?:/i.test(reference),
    )
    const missing = references
      .filter((reference) => !/^https?:/i.test(reference))
      .map((reference) =>
        path.resolve(path.dirname(fixturePath), reference.split(/\s+/)[0]),
      )
      .filter((reference) => !fs.existsSync(reference))
    expect(remote).toEqual([])
    expect(missing).toEqual([])
  })
})
