import { createPptxRenderPlan } from './pptx-plan'
import type { SlideData } from './types'

describe('createPptxRenderPlan', () => {
  it('classifies native, editable, and rasterized slide content', () => {
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [
        {
          url: 'data:image/png;base64,AA==',
          x: 0,
          y: 0,
          width: 1280,
          height: 720,
        },
      ],
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
        {
          type: 'image',
          x: 50,
          y: 100,
          width: 200,
          height: 120,
          src: 'data:image/png;base64,AA==',
          naturalWidth: 200,
          naturalHeight: 120,
          rasterize: true,
        },
      ],
      notes: '',
    }

    const plan = createPptxRenderPlan([slide])

    expect(plan.slides[0].background).toEqual({
      mode: 'native-slide',
      sourceUrl: slide.backgroundImages[0].url,
    })
    expect(plan.slides[0].editableElementCount).toBe(1)
    expect(plan.slides[0].rasterizedElementCount).toBe(1)
    expect(plan.totals.editableAreaRatio).toBeGreaterThan(0)
    expect(plan.totals.editableAreaRatio).toBeLessThan(1)
  })

  it('filters text elements clipped beyond the slide', () => {
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [
        {
          type: 'paragraph',
          x: 0,
          y: 710,
          width: 100,
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

    expect(createPptxRenderPlan([slide]).slides[0].visibleElements).toEqual([])
  })

  it('prunes clipped descendants from visible container elements', () => {
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [
        {
          type: 'container',
          x: 100,
          y: 100,
          width: 500,
          height: 900,
          style: { backgroundColor: 'rgba(0,0,0,0)' },
          children: [
            {
              type: 'paragraph',
              x: 100,
              y: 1000,
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
        },
      ],
      notes: '',
    }

    const plan = createPptxRenderPlan([slide]).slides[0]
    expect(plan.visibleElements[0].type).toBe('container')
    expect(
      plan.visibleElements[0].type === 'container'
        ? plan.visibleElements[0].children
        : null,
    ).toEqual([])
  })

  it('preserves container identity when all descendants remain visible', () => {
    const child = {
      type: 'paragraph' as const,
      x: 100,
      y: 100,
      width: 300,
      height: 30,
      runs: [{ text: 'Label-A' }],
      style: {
        color: 'rgb(0,0,0)',
        fontSize: 16,
        fontFamily: 'Arial',
        fontWeight: 400 as const,
        textAlign: 'left' as const,
        lineHeight: 24,
      },
    }
    const container = {
      type: 'container' as const,
      x: 100,
      y: 100,
      width: 500,
      height: 100,
      style: { backgroundColor: 'rgba(0,0,0,0)' },
      children: [child],
    }
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [container],
      notes: '',
    }

    expect(createPptxRenderPlan([slide]).slides[0].visibleElements[0]).toBe(
      container,
    )
  })

  it('does not count transparent wrapper containers as editable shapes', () => {
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [
        {
          type: 'container',
          x: 0,
          y: 0,
          width: 1000,
          height: 600,
          style: { backgroundColor: 'rgba(0,0,0,0)' },
          children: [
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
        },
      ],
      notes: '',
    }

    const plan = createPptxRenderPlan([slide]).slides[0]
    expect(plan.editableElementCount).toBe(1)
    expect(plan.editableAreaRatio).toBe(1)
  })
})
