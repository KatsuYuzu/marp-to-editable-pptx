import type { RenderSourceRef, TextRun } from './types'

export interface DomSourceTracker {
  sourceFor(element: Element): RenderSourceRef
}

export function createDomSourceTracker(): DomSourceTracker {
  const sourceIds = new WeakMap<Element, string>()
  let nextSourceId = 1

  return {
    sourceFor(element: Element): RenderSourceRef {
      let id = sourceIds.get(element)
      if (!id) {
        id = `node-${nextSourceId++}`
        sourceIds.set(element, id)
      }

      const parent = element.parentElement
      const parentId = parent ? sourceIds.get(parent) : undefined
      const segments: number[] = []
      let current: Element | null = element
      while (current?.parentElement) {
        const siblings = Array.from(current.parentElement.children)
        segments.unshift(siblings.indexOf(current))
        current = current.parentElement
      }
      return {
        id,
        domPath: segments.length > 0 ? `/root/${segments.join('/')}` : '/root',
        ...(parentId ? { parentId } : {}),
      }
    },
  }
}

export function computeAutoScaleFactor(
  runs: TextRun[],
  style: CSSStyleDeclaration,
  rectHeight: number,
  transformScale = 1,
): number {
  const lineHeight = parseFloat(style.lineHeight) || 0
  if (lineHeight <= 0 || rectHeight <= 0) return 1
  const lineCount = runs.filter((run) => run.breakLine).length + 1
  const nonTextHeight =
    (parseFloat(style.paddingTop) || 0) +
    (parseFloat(style.paddingBottom) || 0) +
    (parseFloat(style.borderTopWidth) || 0) +
    (parseFloat(style.borderBottomWidth) || 0)
  const contentHeight = rectHeight - nonTextHeight * transformScale
  if (contentHeight <= 0) return 1
  const naturalHeight = lineCount * lineHeight * transformScale
  const fitScale =
    naturalHeight <= contentHeight * 1.05
      ? 1
      : contentHeight / naturalHeight
  return transformScale * fitScale
}

export function applyAutoScale(
  runs: TextRun[],
  elementFontSize: number,
  scaleFactor: number,
): number {
  if (Math.abs(scaleFactor - 1) < Number.EPSILON) return elementFontSize
  for (const run of runs) {
    if (!run.breakLine && run.fontSize) run.fontSize *= scaleFactor
  }
  return elementFontSize * scaleFactor
}

function parseTransformScale(transform: string | null | undefined): number {
  if (!transform || transform === 'none') return 1
  const values = transform.match(/-?\d*\.?\d+(?:e[-+]?\d+)?/gi)?.map(Number)
  if (!values) return 1
  if (transform.startsWith('scale(')) {
    const scaleX = values[0] ?? 1
    const scaleY = values[1] ?? scaleX
    return Math.sqrt(scaleX * scaleY)
  }

  if (transform.startsWith('matrix3d') && values.length >= 16) {
    const scaleX = Math.hypot(values[0], values[1], values[2])
    const scaleY = Math.hypot(values[4], values[5], values[6])
    return Math.sqrt(scaleX * scaleY)
  }
  if (values.length >= 6) {
    const scaleX = Math.hypot(values[0], values[1])
    const scaleY = Math.hypot(values[2], values[3])
    return Math.sqrt(scaleX * scaleY)
  }
  return 1
}

export function computeElementTransformScale(element: Element): number {
  const view = element.ownerDocument.defaultView
  if (!view) return 1
  let scale = 1
  let current: Element | null = element
  while (current) {
    scale *= parseTransformScale(view.getComputedStyle(current).transform)
    current = current.parentElement
  }
  return Number.isFinite(scale) && scale > 0 ? scale : 1
}
