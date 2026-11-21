import { JSDOM } from 'jsdom'
import {
  applyAutoScale,
  computeAutoScaleFactor,
  createDomSourceTracker,
} from './dom-reader'

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

  it('computes and applies code auto-scale independently', () => {
    const style = {
      lineHeight: '20px',
      paddingTop: '10px',
      paddingBottom: '10px',
    } as CSSStyleDeclaration
    const runs = [
      { text: 'a', fontSize: 24 },
      { text: '', breakLine: true as const },
      { text: 'b', fontSize: 24 },
      { text: '', breakLine: true as const },
      { text: 'c', fontSize: 24 },
    ]

    const scale = computeAutoScaleFactor(runs, style, 50)
    expect(scale).toBe(0.5)
    expect(applyAutoScale(runs, 24, scale)).toBe(12)
    expect(runs.filter((run) => !run.breakLine).map((run) => run.fontSize)).toEqual(
      [12, 12, 12],
    )
  })
})
