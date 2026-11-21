import type { RenderSourceRef } from './types'

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
