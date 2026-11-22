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

  it('does not associate text that the plan clipped from the slide', () => {
    const text = {
      type: 'paragraph' as const,
      x: 50,
      y: 710,
      width: 100,
      height: 8,
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
      x: 40,
      y: 500,
      width: 300,
      height: 220,
      style: { backgroundColor: 'rgb(230,230,230)' },
      children: [],
    }
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [container, text],
      notes: '',
    }

    const plannedSlide = createPptxRenderPlan([slide]).slides[0]
    expect(plannedSlide.containerAssociations).toEqual([])
  })

  it('does not geometrically assign source-owned text to another container', () => {
    const text = {
      type: 'paragraph' as const,
      source: {
        id: 'text-a',
        parentId: 'container-a',
        domPath: '/root/0/0',
      },
      x: 50,
      y: 50,
      width: 100,
      height: 20,
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
    const overlappingContainer = {
      type: 'container' as const,
      source: { id: 'container-b', domPath: '/root/1' },
      x: 40,
      y: 40,
      width: 120,
      height: 40,
      style: { backgroundColor: 'rgb(230,230,230)' },
      children: [],
    }
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [overlappingContainer, text],
      notes: '',
    }

    expect(
      createPptxRenderPlan([slide]).slides[0].containerAssociations,
    ).toEqual([])
  })

  it('lifts explicit z-index children for paint order without mutating source order', () => {
    const ordinary = {
      type: 'container' as const,
      source: { id: 'ordinary', domPath: '/root/0' },
      x: 0,
      y: 200,
      width: 400,
      height: 100,
      style: { backgroundColor: 'rgb(230,241,252)' },
      children: [],
    }
    const overlay = {
      type: 'container' as const,
      source: { id: 'overlay', domPath: '/root/0/0' },
      x: 200,
      y: 220,
      width: 200,
      height: 120,
      stackingPositioned: true,
      stackingIndex: 2,
      style: { backgroundColor: 'rgb(255,255,255)' },
      children: [],
    }
    const host = {
      ...ordinary,
      source: { id: 'host', domPath: '/root/0' },
      children: [overlay],
    }
    const later = {
      ...ordinary,
      source: { id: 'later', domPath: '/root/1' },
      y: 300,
    }
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [host, later],
      notes: '',
    }

    const plan = createPptxRenderPlan([slide]).slides[0]

    expect(slide.elements.map((element) => element.source?.id)).toEqual([
      'host',
      'later',
    ])
    expect(plan.visibleElements.map((element) => element.source?.id)).toEqual([
      'host',
      'later',
      'overlay',
    ])
    expect(
      (plan.visibleElements[0] as any).children,
    ).toEqual([])
  })

  it('does not group text from different DOM owners', () => {
    const base = {
      x: 50,
      width: 100,
      height: 20,
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
    const slide: SlideData = {
      width: 1280,
      height: 720,
      background: 'rgb(255,255,255)',
      backgroundImages: [],
      elements: [
        {
          ...base,
          type: 'paragraph',
          y: 50,
          source: {
            id: 'text-a',
            parentId: 'container-a',
            domPath: '/root/0/0',
          },
        },
        {
          ...base,
          type: 'paragraph',
          y: 72,
          source: {
            id: 'text-b',
            parentId: 'container-b',
            domPath: '/root/1/0',
          },
        },
      ],
      notes: '',
    }

    expect(
      createPptxRenderPlan([slide]).slides[0].elementGroups,
    ).toHaveLength(2)
  })

  it.each(['x', 'y', 'width', 'height'] as const)(
    'rejects invalid %s geometry in a render plan',
    (field) => {
      const slide: SlideData = {
        width: 1280,
        height: 720,
        background: 'rgb(255,255,255)',
        backgroundImages: [],
        elements: [
          {
            type: 'paragraph',
            x: 0,
            y: 0,
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
      ;(slide.elements[0] as any)[field] = Number.NaN

      expect(() => createPptxRenderPlan([slide])).toThrow(
        `invalid ${field} geometry`,
      )
    },
  )
})
