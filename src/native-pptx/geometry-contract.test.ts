import { createPptxRenderPlan } from './pptx-plan'
import type { SlideData, SlideElement } from './types'

function flatten(elements: SlideElement[]): SlideElement[] {
  return elements.flatMap((element) => [
    element,
    ...(element.type === 'container' ? flatten(element.children) : []),
  ])
}

describe('Pptx render plan geometry contract', () => {
  it('keeps visible elements inside the slide canvas', () => {
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [
        {
          type: 'paragraph',
          x: 50,
          y: 50,
          width: 300,
          height: 30,
          runs: [{ text: 'Label-A' }],
          style: {
            color: 'rgb(0,0,0)',
            fontSize: 16,
            fontFamily: 'Arial',
            fontWeight: 400,
            textAlign: 'left',
            lineHeight: 24,
          },
        },
      ],
      notes: '',
    }

    const plan = createPptxRenderPlan([slide])
    for (const element of flatten(plan.slides[0].visibleElements)) {
      if (element.type === 'image') continue
      expect(element.x).toBeGreaterThanOrEqual(-1)
      expect(element.y).toBeGreaterThanOrEqual(-1)
      expect(element.x).toBeLessThanOrEqual(slide.width)
      expect(element.y).toBeLessThanOrEqual(slide.height)
    }
  })

  it('reports deterministic editable and raster metrics', () => {
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [
        {
          type: 'image',
          x: 0,
          y: 0,
          width: 100,
          height: 100,
          src: 'data:image/png;base64,AA==',
          naturalWidth: 100,
          naturalHeight: 100,
          rasterize: true,
        },
      ],
      notes: '',
    }

    const first = createPptxRenderPlan([slide])
    const second = createPptxRenderPlan([slide])
    expect(second.totals).toEqual(first.totals)
    expect(first.slides[0].rasterizedElementCount).toBe(1)
  })

  it.each([
    [960, 720],
    [1080, 1920],
  ])('preserves a %sx%s source canvas without 16:9 assumptions', (
    width,
    height,
  ) => {
    const slide: SlideData = {
      width,
      height,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [{
        type: 'paragraph',
        x: 20,
        y: 20,
        width: width - 40,
        height: 30,
        runs: [{ text: 'Label-A' }],
        style: {
          color: 'rgb(0,0,0)',
          fontSize: 16,
          fontFamily: 'Arial',
          fontWeight: 400,
          textAlign: 'left',
          lineHeight: 24,
        },
      }],
      notes: '',
    }

    const plan = createPptxRenderPlan([slide])
    expect(plan.slides[0].source.width).toBe(width)
    expect(plan.slides[0].source.height).toBe(height)
    expect(plan.slides[0].visibleElements[0].x).toBeLessThan(width)
    expect(plan.slides[0].visibleElements[0].y).toBeLessThan(height)
  })
})
