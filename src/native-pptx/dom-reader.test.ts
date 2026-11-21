import { JSDOM } from 'jsdom'
import { createDomSourceTracker } from './dom-reader'

describe('createDomSourceTracker', () => {
  it('returns stable IDs and source DOM paths', () => {
    const dom = new JSDOM(
      '<!doctype html><html><body><div><span></span></div></body></html>',
    )
    const parent = dom.window.document.querySelector('div')!
    const child = dom.window.document.querySelector('span')!
    const tracker = createDomSourceTracker()

    const parentSource = tracker.sourceFor(parent)
    const childSource = tracker.sourceFor(child)

    expect(tracker.sourceFor(parent)).toEqual(parentSource)
    expect(childSource.parentId).toBe(parentSource.id)
    expect(childSource.domPath).not.toBe(parentSource.domPath)
    dom.window.close()
  })
})
