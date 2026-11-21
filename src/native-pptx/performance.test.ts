import { performance } from 'node:perf_hooks'
import { createPptxRenderPlan } from './pptx-plan'
import type { SlideData } from './types'

describe('Pptx render plan performance', () => {
  it('plans a 200-slide deck within the deterministic budget', () => {
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: Array.from({ length: 100 }, (_, index) => ({
        type: 'paragraph' as const,
        x: 50,
        y: 20 + index * 5,
        width: 300,
        height: 20,
        runs: [{ text: `Label-${index}` }],
        style: {
          color: 'rgb(0,0,0)',
          fontSize: 16,
          fontFamily: 'Arial',
          fontWeight: 400,
          textAlign: 'left' as const,
          lineHeight: 24,
        },
      })),
      notes: '',
    }
    const slides = Array.from({ length: 200 }, () => ({
      ...slide,
      elements: slide.elements.map((element) => ({ ...element })),
    }))

    const started = performance.now()
    const plan = createPptxRenderPlan(slides)
    const elapsed = performance.now() - started

    expect(plan.slides).toHaveLength(200)
    expect(elapsed).toBeLessThan(250)
  })
})
