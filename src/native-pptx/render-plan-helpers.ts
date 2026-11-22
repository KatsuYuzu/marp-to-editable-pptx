import type { ContainerElement, SlideElement } from './types'
import { isTransparent } from './utils'

const EMBEDDABLE_IN_SHAPE = new Set([
  'paragraph',
  'heading',
  'list',
  'blockquote',
])
const GROUPABLE_TYPES = new Set(['paragraph', 'blockquote'])
const MAX_GROUP_GAP_PX = 64
const CONTAINMENT_SLOP_PX = 4
const MAX_X_DIFF_PX = 5
const MAX_WIDTH_DIFF_PX = 5

export function isSimpleTextContainer(element: ContainerElement): boolean {
  if (!element.children || element.children.length === 0) return false
  if (element.runs && element.runs.length > 0) return false
  if (element.style.borderBottom) return false
  if (!element.children.every((child) => EMBEDDABLE_IN_SHAPE.has(child.type)))
    return false
  if (element.children.length > 1) {
    const xs = element.children.map((child) => child.x)
    if (Math.max(...xs) - Math.min(...xs) > 50) return false
  }
  return true
}

function isContainedIn(text: SlideElement, container: SlideElement): boolean {
  const strict =
    text.x >= container.x - CONTAINMENT_SLOP_PX &&
    text.y >= container.y - CONTAINMENT_SLOP_PX &&
    text.x + text.width <=
      container.x + container.width + CONTAINMENT_SLOP_PX &&
    text.y + text.height <=
      container.y + container.height + CONTAINMENT_SLOP_PX
  if (strict) return true

  const dx = Math.abs(text.x - container.x)
  const dy = Math.abs(text.y - container.y)
  const dw = Math.abs(text.width - container.width)
  const dh = Math.abs(text.height - container.height)
  return dx <= 16 && dy <= 16 && dw <= 16 && dh <= 16
}

function sourceOwns(
  text: SlideElement,
  container: ContainerElement,
): boolean {
  if (!text.source || !container.source) return false
  return (
    text.source.id === container.source.id ||
    text.source.parentId === container.source.id ||
    text.source.domPath.startsWith(`${container.source.domPath}/`)
  )
}

function hasKnownDifferentOwner(
  text: SlideElement,
  container: ContainerElement,
): boolean {
  return !!(text.source && container.source) && !sourceOwns(text, container)
}

function isEmbeddableContainer(
  element: ContainerElement,
  slideW: number,
  slideH: number,
): boolean {
  if (element.runs && element.runs.length > 0) return false
  if (element.style.borderBottom) return false
  if (
    element.children.length > 0 &&
    !isSimpleTextContainer(element)
  ) {
    return false
  }
  const hasBackground = !isTransparent(element.style.backgroundColor)
  const hasBorderLeft =
    !!element.style.borderLeft && element.style.borderLeft.width > 0
  const hasBorder =
    Object.keys(element.style.borders ?? {}).length > 0 ||
    ((element.style.borderWidth ?? 0) > 0 &&
      !!element.style.borderColor &&
      !isTransparent(element.style.borderColor))
  const hasShadow = element.style.boxShadow === true
  if (!hasBackground && !hasBorder && !hasShadow && !hasBorderLeft) return false
  return !(
    slideW > 0 &&
    slideH > 0 &&
    element.width >= slideW * 0.9 &&
    element.height >= slideH * 0.9
  )
}

export function associateContainerText(
  elements: SlideElement[],
  slideW: number,
  slideH: number,
): Map<ContainerElement, SlideElement[]> {
  const result = new Map<ContainerElement, SlideElement[]>()
  const containers = elements
    .filter(
      (element): element is ContainerElement =>
        element.type === 'container' &&
        isEmbeddableContainer(element, slideW, slideH),
    )
    .sort((left, right) => left.width * left.height - right.width * right.height)
  const textElements = elements.filter((element) =>
    EMBEDDABLE_IN_SHAPE.has(element.type),
  )
  const assigned = new Set<SlideElement>()

  for (const container of containers) {
    const owned = textElements
      .filter(
        (element) =>
          !assigned.has(element) &&
          sourceOwns(element, container),
      )
      .sort((left, right) => left.y - right.y)
    if (owned.length > 0) {
      result.set(container, owned)
      owned.forEach((element) => assigned.add(element))
    }
  }

  for (const container of containers) {
    if (result.has(container)) continue
    const inside = textElements
      .filter(
        (element) =>
          !assigned.has(element) &&
          !hasKnownDifferentOwner(element, container) &&
          isContainedIn(element, container),
      )
      .sort((left, right) => left.y - right.y)
    if (inside.length > 0) {
      result.set(container, inside)
      inside.forEach((element) => assigned.add(element))
    }
  }

  return result
}

function sourceOwnerKey(element: SlideElement): string | undefined {
  if (!element.source) return undefined
  return element.source.parentId ?? element.source.domPath
}

function canMergeElements(left: SlideElement, right: SlideElement): boolean {
  if (!GROUPABLE_TYPES.has(left.type) || !GROUPABLE_TYPES.has(right.type))
    return false
  const leftOwner = sourceOwnerKey(left)
  const rightOwner = sourceOwnerKey(right)
  if (leftOwner && rightOwner && leftOwner !== rightOwner) return false
  for (const element of [left, right]) {
    if (
      element.type === 'blockquote' &&
      element.borderLeft &&
      element.borderLeft.width > 0
    )
      return false
  }
  for (const element of [left, right]) {
    if (element.type !== 'paragraph') continue
    if (
      element.style.lineHeight > 0 &&
      element.height > element.style.lineHeight * 1.5
    )
      return false
  }
  if (Math.abs(left.x - right.x) > MAX_X_DIFF_PX) return false
  if (Math.abs(left.width - right.width) > MAX_WIDTH_DIFF_PX) return false
  const gap = right.y - (left.y + left.height)
  if (gap < 0 || gap > MAX_GROUP_GAP_PX) return false
  return true
}

export function groupAdjacentTextElements(
  elements: SlideElement[],
): SlideElement[][] {
  if (elements.length === 0) return []
  const groups: SlideElement[][] = [[elements[0]]]
  for (const current of elements.slice(1)) {
    const lastGroup = groups[groups.length - 1]
    const previous = lastGroup[lastGroup.length - 1]
    if (canMergeElements(previous, current)) lastGroup.push(current)
    else groups.push([current])
  }
  return groups
}
