import { JSDOM } from 'jsdom'
import {
  applyAutoScale,
  computeAutoScaleFactor,
  computeElementTransformScale,
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

  it('computes transformed font scale in the post-transform coordinate system', () => {
    const dom = new JSDOM(
      '<!doctype html><html><body><div style="transform:scale(0.5)"><pre style="padding:10px;line-height:30px"></pre></div></body></html>',
    )
    const pre = dom.window.document.querySelector('pre')!
    Object.defineProperty(pre, 'getBoundingClientRect', {
      value: () => ({
        height: 55,
      }),
    })
    const style = dom.window.getComputedStyle(pre)
    const transformScale = computeElementTransformScale(pre)
    const runs = [
      { text: 'Label-A' },
      { text: '', breakLine: true as const },
      { text: 'Label-B' },
      { text: '', breakLine: true as const },
      { text: 'Label-C' },
    ]

    expect(transformScale).toBeCloseTo(0.5)
    expect(
      computeAutoScaleFactor(
        runs,
        style,
        pre.getBoundingClientRect().height,
        transformScale,
      ),
    ).toBeCloseTo(0.5)
    dom.window.close()
  })

  it('follows transform scale across a Shadow DOM host boundary', () => {
    const dom = new JSDOM(
      '<!doctype html><html><body><div id="host"></div></body></html>',
    )
    const host = dom.window.document.querySelector('#host')! as HTMLElement
    host.style.transform = 'scale(0.5)'
    const shadow = host.attachShadow({ mode: 'open' })
    const child = dom.window.document.createElement('pre')
    shadow.append(child)
    const OriginalShadowRoot = globalThis.ShadowRoot
    ;(globalThis as any).ShadowRoot = dom.window.ShadowRoot

    expect(computeElementTransformScale(child)).toBeCloseTo(0.5)

    ;(globalThis as any).ShadowRoot = OriginalShadowRoot
    dom.window.close()
  })

  it('does not count a trailing newline as a rendered code line', () => {
    const style = {
      lineHeight: '28.3475px',
      paddingTop: '16px',
      paddingBottom: '16px',
    } as CSSStyleDeclaration
    const runs = [
      { text: '.cat { color: blue; }', fontSize: 24.65 },
      { text: '', breakLine: true as const },
    ]

    expect(computeAutoScaleFactor(runs, style, 63)).toBe(1)
  })
})
