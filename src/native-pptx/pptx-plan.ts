import type { SlideData, SlideElement } from './types'
import { isTransparent, rgbToHex } from './utils'

export type BackgroundPlan =
  | { mode: 'native-slide'; sourceUrl: string }
  | { mode: 'shapes' }

export interface SlidePptxPlan {
  source: SlideData
  background: BackgroundPlan
  visibleElements: SlideElement[]
  visualBgMayBeDark: boolean
  editableElementCount: number
  rasterizedElementCount: number
  editableAreaRatio: number
}

export interface PptxRenderPlan {
  slides: SlidePptxPlan[]
  totals: {
    editableElementCount: number
    rasterizedElementCount: number
    editableAreaRatio: number
  }
}

const TEXT_ELEMENT_TYPES = new Set([
  'heading',
  'paragraph',
  'list',
  'table',
  'code',
  'blockquote',
  'header',
  'footer',
])

function visibleElement(
  element: SlideElement,
  slideHeight: number,
): SlideElement | undefined {
  if (TEXT_ELEMENT_TYPES.has(element.type) && slideHeight - element.y < 20) {
    return undefined
  }
  if (element.type !== 'container') return element
  const children = element.children
    .map((child) => visibleElement(child, slideHeight))
    .filter((child): child is SlideElement => !!child)
  if (
    children.length === element.children.length &&
    children.every((child, index) => child === element.children[index])
  ) {
    return element
  }
  return { ...element, children }
}

function hasVisibleContainerShape(element: SlideElement): boolean {
  if (element.type !== 'container') return true
  const style = element.style
  return (
    !isTransparent(style.backgroundColor) ||
    (style.borderWidth ?? 0) > 0 ||
    style.borderLeft !== undefined ||
    style.borderBottom !== undefined ||
    style.boxShadow === true
  )
}

function classifyElements(elements: SlideElement[]): {
  nativeCount: number
  rasterCount: number
  nativeArea: number
  totalArea: number
} {
  const result = { nativeCount: 0, rasterCount: 0, nativeArea: 0, totalArea: 0 }
  for (const element of elements) {
    const ownArea = Math.max(0, element.width) * Math.max(0, element.height)
    const isRaster = element.type === 'image' && element.rasterize === true
    const ownsNativeShape = !isRaster && hasVisibleContainerShape(element)

    if (isRaster) result.rasterCount++
    else if (ownsNativeShape) {
      result.nativeCount++
      result.nativeArea += ownArea
    }
    if (ownsNativeShape || isRaster) result.totalArea += ownArea

    if (element.type === 'container') {
      const children = classifyElements(element.children)
      result.nativeCount += children.nativeCount
      result.rasterCount += children.rasterCount
      result.nativeArea += children.nativeArea
      result.totalArea += children.totalArea
    }
  }
  return result
}

export function createPptxRenderPlan(slides: SlideData[]): PptxRenderPlan {
  const plannedSlides = slides.map((slide): SlidePptxPlan => {
    const backgrounds = slide.backgroundImages ?? []
    const firstBackground = backgrounds[0]
    const isFullSlide =
      firstBackground &&
      !firstBackground.cssFilter &&
      !firstBackground.backgroundSizeContain &&
      firstBackground.x <= 1 &&
      firstBackground.y <= 1 &&
      Math.abs(firstBackground.width - slide.width) <= 2 &&
      Math.abs(firstBackground.height - slide.height) <= 2

    const visibleElements = slide.elements
      .map((element) => visibleElement(element, slide.height))
      .filter((element): element is SlideElement => !!element)
    const classification = classifyElements(visibleElements)
    const cssIsFallbackWhite =
      isTransparent(slide.background) ||
      rgbToHex(slide.background).toUpperCase() === 'FFFFFF'
    const visualBgMayBeDark =
      backgrounds.some(
        (background) =>
          background.url !== '' &&
          !background.fromCssFallback &&
          !background.backgroundSizeContain &&
          background.width >= slide.width * 0.8,
      ) && cssIsFallbackWhite

    return {
      source: slide,
      background:
        isFullSlide && backgrounds.length === 1
          ? { mode: 'native-slide', sourceUrl: firstBackground.url }
          : { mode: 'shapes' },
      visibleElements,
      visualBgMayBeDark,
      editableElementCount: classification.nativeCount,
      rasterizedElementCount: classification.rasterCount,
      editableAreaRatio:
        classification.totalArea > 0
          ? Math.min(1, classification.nativeArea / classification.totalArea)
          : 1,
    }
  })

  const editableElementCount = plannedSlides.reduce(
    (sum, slide) => sum + slide.editableElementCount,
    0,
  )
  const rasterizedElementCount = plannedSlides.reduce(
    (sum, slide) => sum + slide.rasterizedElementCount,
    0,
  )
  const totals = plannedSlides.reduce(
    (aggregate, slide) => {
      const classification = classifyElements(slide.visibleElements)
      aggregate.editableArea += classification.nativeArea
      aggregate.totalArea += classification.totalArea
      return aggregate
    },
    { editableArea: 0, totalArea: 0 },
  )

  return {
    slides: plannedSlides,
    totals: {
      editableElementCount,
      rasterizedElementCount,
      editableAreaRatio:
        totals.totalArea > 0 ? totals.editableArea / totals.totalArea : 1,
    },
  }
}
