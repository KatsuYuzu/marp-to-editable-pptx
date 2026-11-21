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
): number {
  const lineHeight = parseFloat(style.lineHeight) || 0
  if (lineHeight <= 0 || rectHeight <= 0) return 1
  const lineCount = runs.filter((run) => run.breakLine).length + 1
  const padding =
    (parseFloat(style.paddingTop) || 0) +
    (parseFloat(style.paddingBottom) || 0)
  const contentHeight = rectHeight - padding
  if (contentHeight <= 0) return 1
  const naturalHeight = lineCount * lineHeight
  return naturalHeight <= contentHeight * 1.05
    ? 1
    : contentHeight / naturalHeight
}

export function applyAutoScale(
  runs: TextRun[],
  elementFontSize: number,
  scaleFactor: number,
): number {
  if (scaleFactor >= 1) return elementFontSize
  for (const run of runs) {
    if (!run.breakLine && run.fontSize) run.fontSize *= scaleFactor
  }
  return elementFontSize * scaleFactor
}
