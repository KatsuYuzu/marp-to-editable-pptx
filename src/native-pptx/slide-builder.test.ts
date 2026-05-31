import {
  buildPptx,
  placeElement,
  toTextProps,
  toListTextProps,
  groupAdjacentTextElements,
  associateContainerText,
} from './slide-builder'
import type { SlideData, ImageElement, SlideElement, TableCell, TableCellParagraph } from './types'

// pptxgenjs creates a real object; we spy on its methods to verify calls.
// We do NOT jest.mock('pptxgenjs') so that buildPptx() internally
// instantiates a real PptxGenJS instance whose prototype methods we can spy on.

describe('buildPptx', () => {
  const minimalSlide: SlideData = {
    width: 1280,
    height: 720,
    background: 'rgb(255, 255, 255)',
    backgroundImages: [],
    elements: [],
    notes: '',
  }

  it('defines layout with Marp default size (1280x720)', () => {
    const pptx = buildPptx([minimalSlide])
    // pptxgenjs stores layout name
    expect(pptx.layout).toBe('MARP')
  })

  it('calls addSlide for each slide', () => {
    const pptx = buildPptx([minimalSlide, minimalSlide])
    // PptxGenJS stores slides internally; we verify by checking the count
    // through its write path or by accessing internal state.
    // Since pptxgenjs doesn't expose a slide count getter, just verify
    // no error is thrown and the object is returned.
    expect(pptx).toBeDefined()
  })

  it('sets slide background color', () => {
    const pptx = buildPptx([minimalSlide])
    // Verify through the generated PPTX - we rely on pptxgenjs internals.
    // A basic smoke test: the pptx object should be writable.
    expect(typeof pptx.write).toBe('function')
  })

  it('sets presenter notes', () => {
    const slideWithNotes: SlideData = {
      ...minimalSlide,
      notes: 'These are presenter notes',
    }
    // Smoke test: no errors thrown
    const pptx = buildPptx([slideWithNotes])
    expect(pptx).toBeDefined()
  })

  it('does not add PowerPoint slide numbers when the source deck has no pagination metadata', () => {
    const pptx = buildPptx([minimalSlide])
    const internalSlides = (pptx as any)._slides as any[]
    expect(internalSlides).toHaveLength(1)
    expect(internalSlides[0]._slideNumberProps).toBeFalsy()
  })

  it('adds PowerPoint slide numbers to every slide when the source deck uses pagination', () => {
    const pptx = buildPptx([
      { ...minimalSlide, sourceHasPagination: true },
      { ...minimalSlide },
    ])
    const internalSlides = (pptx as any)._slides as any[]
    expect(internalSlides).toHaveLength(2)
    expect(internalSlides[0]._slideNumberProps).toBeDefined()
    expect(internalSlides[1]._slideNumberProps).toBeDefined()
    expect(internalSlides[0]._slideNumberProps).toMatchObject({
      align: 'right',
      color: '777777',
    })
    expect(internalSlides[1]._slideNumberProps).toMatchObject({
      align: 'right',
      color: '777777',
    })
  })

  it('uses a fixed deck-wide placement and styling for the native slide number field', () => {
    const pptx = buildPptx([
      {
        ...minimalSlide,
        sourceHasPagination: true,
      },
    ])

    const internalSlides = (pptx as any)._slides as any[]
    expect(internalSlides).toHaveLength(1)
    expect(internalSlides[0]._slideNumberProps).toMatchObject({
      align: 'right',
      color: '777777',
    })
    expect(internalSlides[0]._slideNumberProps.x).toBeCloseTo(0, 4)
    expect(internalSlides[0]._slideNumberProps.y).toBeCloseTo(666 / 96, 4)
    expect(internalSlides[0]._slideNumberProps.w).toBeCloseTo(1240 / 96, 4)
    expect(internalSlides[0]._slideNumberProps.fontSize).toBeCloseTo(13.5, 4)
  })

  it('places heading and paragraph elements without error', () => {
    const slideWithElements: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'heading',
          level: 1,
          runs: [
            {
              text: 'Title',
              color: 'rgb(34, 68, 102)',
              fontSize: 40,
              fontFamily: '"Noto Sans JP"',
              bold: true,
            },
          ],
          x: 70,
          y: 80,
          width: 1140,
          height: 60,
          style: {
            color: 'rgb(34, 68, 102)',
            fontSize: 40,
            fontFamily: '"Noto Sans JP"',
            fontWeight: 700,
            textAlign: 'left',
            lineHeight: 48,
          },
        },
        {
          type: 'paragraph',
          runs: [
            {
              text: 'Body text',
              color: 'rgb(51, 51, 51)',
              fontSize: 24,
              fontFamily: 'Arial',
              bold: false,
            },
          ],
          x: 70,
          y: 160,
          width: 1140,
          height: 30,
          style: {
            color: 'rgb(51, 51, 51)',
            fontSize: 24,
            fontFamily: 'Arial',
            fontWeight: 400,
            textAlign: 'left',
            lineHeight: 36,
          },
        },
      ],
    }
    const pptx = buildPptx([slideWithElements])
    expect(pptx).toBeDefined()
  })

  it('places list elements without error', () => {
    const slideWithList: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'list',
          ordered: false,
          items: [
            {
              text: 'Item 1',
              level: 0,
              runs: [{ text: 'Item 1', fontSize: 18 }],
            },
            {
              text: 'Item 2',
              level: 0,
              runs: [{ text: 'Item 2', fontSize: 18 }],
            },
            {
              text: 'Nested item',
              level: 1,
              runs: [{ text: 'Nested item', fontSize: 16 }],
            },
          ],
          x: 70,
          y: 200,
          width: 600,
          height: 120,
          style: {
            color: 'rgb(0, 0, 0)',
            fontSize: 18,
            fontFamily: 'Arial',
            fontWeight: 400,
            textAlign: 'left',
            lineHeight: 27,
          },
        },
      ],
    }
    const pptx = buildPptx([slideWithList])
    expect(pptx).toBeDefined()
  })

  it('places container elements with children without error', () => {
    const slideWithContainer: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'container',
          children: [
            {
              type: 'paragraph',
              runs: [{ text: 'Nested paragraph' }],
              x: 80,
              y: 90,
              width: 500,
              height: 24,
              style: {
                color: 'rgb(0, 0, 0)',
                fontSize: 16,
                fontFamily: 'Arial',
                fontWeight: 400,
                textAlign: 'left',
                lineHeight: 24,
              },
            },
          ],
          x: 70,
          y: 80,
          width: 640,
          height: 200,
          style: { backgroundColor: 'rgb(240, 240, 240)' },
        },
      ],
    }
    const pptx = buildPptx([slideWithContainer])
    expect(pptx).toBeDefined()
  })

  it('places bordered container without error', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'container',
          children: [
            {
              type: 'paragraph',
              runs: [{ text: 'Card text' }],
              x: 80,
              y: 90,
              width: 480,
              height: 24,
              style: {
                color: 'rgb(0, 0, 0)',
                fontSize: 16,
                fontFamily: 'Arial',
                fontWeight: 400,
                textAlign: 'left',
                lineHeight: 24,
              },
            },
          ],
          x: 70,
          y: 80,
          width: 500,
          height: 100,
          style: {
            backgroundColor: 'rgb(255, 244, 232)',
            borderWidth: 1,
            borderColor: 'rgb(207, 216, 227)',
            borderRadius: 12,
          },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })

  it('places border-only container without background', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'container',
          children: [
            {
              type: 'paragraph',
              runs: [{ text: 'Text' }],
              x: 80,
              y: 90,
              width: 480,
              height: 24,
              style: {
                color: 'rgb(0, 0, 0)',
                fontSize: 16,
                fontFamily: 'Arial',
                fontWeight: 400,
                textAlign: 'left',
                lineHeight: 24,
              },
            },
          ],
          x: 70,
          y: 80,
          width: 500,
          height: 100,
          style: {
            backgroundColor: 'rgba(0, 0, 0, 0)',
            borderWidth: 1,
            borderColor: 'rgb(207, 216, 227)',
          },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })
})

describe('toTextProps', () => {
  it('converts TextRun to pptxgenjs TextProps', () => {
    const result = toTextProps({
      text: 'Hello',
      color: 'rgb(255, 0, 0)',
      fontSize: 24,
      fontFamily: '"Noto Sans JP", sans-serif',
      bold: true,
      italic: false,
    })

    expect(result.text).toBe('Hello')
    expect(result.options?.color).toBe('FF0000')
    expect(result.options?.fontSize).toBe(18) // 24 * 0.75
    expect(result.options?.fontFace).toBe('Noto Sans JP')
    expect(result.options?.bold).toBe(true)
  })

  it('converts hyperlinks', () => {
    const result = toTextProps({
      text: 'Link',
      hyperlink: 'https://example.com',
    })
    expect(result.options?.hyperlink).toEqual({ url: 'https://example.com' })
  })

  it('converts breakLine:true run to breakLine:true TextProps', () => {
    const result = toTextProps({ text: '', breakLine: true })
    expect(result.text).toBe('')
    expect(result.options?.breakLine).toBe(true)
    // Should not carry color or font overrides
    expect(result.options?.color).toBeUndefined()
  })

  it('converts backgroundColor to 6-digit hex highlight', () => {
    const result = toTextProps({
      text: 'Highlighted',
      color: 'rgb(0, 0, 0)',
      fontSize: 22,
      backgroundColor: 'rgb(241, 196, 15)',
    })
    expect(result.text).toBe('Highlighted')
    expect(result.options?.highlight).toBe('F1C40F')
  })

  it('omits highlight when backgroundColor is absent', () => {
    const result = toTextProps({
      text: 'Normal',
      color: 'rgb(0, 0, 0)',
      fontSize: 16,
    })
    expect(result.options?.highlight).toBeUndefined()
  })

  it('preserves light-gray highlight for semi-transparent rgba (Marp inline <code> pattern)', () => {
    // rgba(129, 139, 152, 0.12) is the actual computed backgroundColor for
    // Marp default theme inline <code> elements (verified via MARP_PPTX_DEBUG).
    // Without alpha compositing, rgbToHex strips alpha and returns #818B98
    // (opaque medium grey), which PowerPoint renders as a visibly dark highlight.
    // After compositing over white: rgb(240, 241, 243) — channels 240–243 ≤ 248
    // → highlight is preserved as #F0F1F3 (subtle light-gray, better than dark).
    const result = toTextProps({
      text: 'inline code',
      color: 'rgb(0, 0, 0)',
      fontSize: 16,
      backgroundColor: 'rgba(129, 139, 152, 0.12)',
    })
    expect(result.options?.highlight).toBe('F0F1F3')
  })

  it('preserves light-gray highlight for rgba(0,0,0,0.06) (faint dark-over-white code bg)', () => {
    // rgba(0,0,0,0.06) composited over white = rgb(240,240,240) — channels 240 ≤ 248
    // → highlight preserved as #F0F0F0 (subtle light-gray).
    const result = toTextProps({
      text: 'code',
      color: 'rgb(0, 0, 0)',
      fontSize: 14,
      backgroundColor: 'rgba(0, 0, 0, 0.06)',
    })
    expect(result.options?.highlight).toBe('F0F0F0')
  })

  it('omits highlight for near-pure-white rgba (essentially invisible)', () => {
    // rgba(0,0,0,0.02) composited = rgb(250,250,250) — all channels 250 > 248 → suppressed.
    const result = toTextProps({
      text: 'ghost',
      color: 'rgb(0, 0, 0)',
      fontSize: 14,
      backgroundColor: 'rgba(0, 0, 0, 0.02)',
    })
    expect(result.options?.highlight).toBeUndefined()
  })

  it('suppresses light-gray highlight when text color is also light (image-backed dark slide)', () => {
    // rgba(129,139,152,0.12) composited over white = #F0F1F3 (near-white).
    // If the text is also white (dark-background slide where CSS bg is still white
    // because the darkness comes from a bg image), applying #F0F1F3 highlight
    // would hide white text. Both highlight and text are "light" (all ch > 200) → suppress.
    const result = toTextProps(
      {
        text: 'inline code',
        color: 'rgb(255, 255, 255)', // white text (dark slide)
        fontSize: 16,
        backgroundColor: 'rgba(129, 139, 152, 0.12)',
      },
      // slideBg = white (image-backed dark slide: CSS bg-color is still rgb(255,255,255))
      'rgb(255, 255, 255)',
    )
    expect(result.options?.highlight).toBeUndefined()
  })

  it('composites rgba over actual dark CSS bg and keeps visible highlight', () => {
    // On a CSS-dark slide (background-color set to dark), compositing gives correct dark result.
    // rgba(129,139,152,0.12) over rgb(30,30,36):
    //   r = 30 + (129-30)*0.12 ≈ 42
    //   g = 30 + (139-30)*0.12 ≈ 43
    //   b = 36 + (152-36)*0.12 ≈ 50
    // delta from bg: max(12,13,14) = 14 ≥ 10 (lowered threshold) → kept
    const result = toTextProps(
      {
        text: 'inline code',
        color: 'rgb(255, 255, 255)',
        fontSize: 16,
        backgroundColor: 'rgba(129, 139, 152, 0.12)',
      },
      'rgb(30, 30, 36)', // actual CSS dark bg
    )
    // delta = 14 ≥ 10 threshold → highlight is now preserved (visible subtle tint)
    expect(result.options?.highlight).toBeDefined()
  })

  it('composites rgba over actual dark CSS bg and shows highlight when contrast is sufficient', () => {
    // Strong highlight rgba(100,200,100,0.5) over dark bg rgb(30,30,36):
    //   r = 30 + (100-30)*0.5 = 65 → delta from bg = 35 ≥ 15 → kept
    const result = toTextProps(
      {
        text: 'highlighted',
        color: 'rgb(255, 255, 255)',
        fontSize: 16,
        backgroundColor: 'rgba(100, 200, 100, 0.5)',
      },
      'rgb(30, 30, 36)',
    )
    expect(result.options?.highlight).toBeDefined()
  })

  it('keeps highlight for yellow marker even when text is light', () => {
    // Yellow marker #FFF2A8: composited rgb(255,243,178), b=178 ≤ 200 → NOT all-light → kept
    // even with white text, because the blue channel 178 < 200 breaks the all-light check.
    const result = toTextProps({
      text: 'marked',
      color: 'rgb(255, 255, 255)', // white text
      fontSize: 16,
      backgroundColor: 'rgba(255, 242, 168, 0.9)',
    })
    expect(result.options?.highlight).toBeDefined()
  })

  it('suppresses light highlight when visualBgMayBeDark=true, even if text is not pure white', () => {
    // Scenario: image-backed dark slide.  CSS bg = white (fallback), but visual bg is dark.
    // Code text color is a Marp-theme grayish-light, NOT pure white (r=210).
    // Rule 4 (text-lightness) still fires (all ch > 200), but this tests that the
    // 3rd argument (visualBgMayBeDark=true) alone would also suppress it via rule 3.
    const result = toTextProps(
      {
        text: 'code',
        color: 'rgb(210, 215, 220)', // light but not pure white
        fontSize: 16,
        backgroundColor: 'rgba(129, 139, 152, 0.12)',
      },
      'rgb(255, 255, 255)', // CSS bg = white fallback
      true, // visualBgMayBeDark
    )
    expect(result.options?.highlight).toBeUndefined()
  })

  it('keeps highlight when visualBgMayBeDark=false and text is dark (slide 42 case)', () => {
    // White bg, no bg images → visualBgMayBeDark=false.
    // rgba(0.12) → #F0F1F3, delta=15 from white → NOT < 15 → kept.
    // Text is dark so text-lightness check doesn't fire.
    const result = toTextProps(
      {
        text: 'code',
        color: 'rgb(51, 51, 51)', // typical dark-on-white text
        fontSize: 16,
        backgroundColor: 'rgba(129, 139, 152, 0.12)',
      },
      'rgb(255, 255, 255)',
      false, // visualBgMayBeDark = false (slide 42 case)
    )
    expect(result.options?.highlight).toBe('F0F1F3')
  })

  it('keeps highlight for clearly saturated rgba (yellow marker)', () => {
    // rgba(255, 242, 168, 0.9) → composited rgb(255, 243, 178) → g=243 ≤ 248 → kept
    const result = toTextProps({
      text: 'marked',
      color: 'rgb(0, 0, 0)',
      fontSize: 16,
      backgroundColor: 'rgba(255, 242, 168, 0.9)',
    })
    expect(result.options?.highlight).toBeDefined()
    expect(result.options?.highlight).not.toBe(undefined)
  })

  it('passes subscript:true to PptxGenJS options', () => {
    const result = toTextProps({
      text: '2',
      color: 'rgb(0, 0, 0)',
      fontSize: 12,
      subscript: true,
    })
    expect(result.options?.subscript).toBe(true)
    expect(result.options?.superscript).toBeUndefined()
  })

  it('passes superscript:true to PptxGenJS options', () => {
    const result = toTextProps({
      text: '2',
      color: 'rgb(0, 0, 0)',
      fontSize: 12,
      superscript: true,
    })
    expect(result.options?.superscript).toBe(true)
    expect(result.options?.subscript).toBeUndefined()
  })

  it('does not set subscript/superscript when both are absent', () => {
    const result = toTextProps({
      text: 'normal',
      color: 'rgb(0, 0, 0)',
      fontSize: 16,
    })
    expect(result.options?.subscript).toBeUndefined()
    expect(result.options?.superscript).toBeUndefined()
  })
})

describe('toListTextProps', () => {
  it('converts list item to TextProps with bullet', () => {
    const result = toListTextProps({
      text: 'Item',
      level: 0,
      runs: [
        {
          text: 'Item',
          color: 'rgb(0, 0, 0)',
          fontSize: 16,
          fontFamily: 'Arial',
        },
      ],
    })

    expect(result).toHaveLength(1)
    expect(result[0].options?.bullet).toBe(true)
    expect(result[0].options?.indentLevel).toBe(0)
  })

  it('converts nesting level to indentLevel', () => {
    const result = toListTextProps({
      text: 'Nested',
      level: 2,
      runs: [{ text: 'Nested', fontSize: 14 }],
    })

    expect(result[0].options?.indentLevel).toBe(2)
  })

  it('falls back to text when runs is empty', () => {
    const result = toListTextProps({
      text: 'Fallback',
      level: 0,
      runs: [],
    })

    expect(result).toHaveLength(1)
    expect(result[0].text).toBe('Fallback')
    expect(result[0].options?.bullet).toBe(true)
  })

  it('uses numbered bullet when ordered=true', () => {
    const result = toListTextProps(
      {
        text: 'Numbered',
        level: 0,
        runs: [{ text: 'Numbered', fontSize: 16 }],
      },
      true,
    )

    expect(result[0].options?.bullet).toEqual({
      type: 'number',
      style: 'arabicPeriod',
    })
  })

  it('uses plain bullet when ordered=false', () => {
    const result = toListTextProps(
      {
        text: 'Bullet',
        level: 0,
        runs: [{ text: 'Bullet', fontSize: 16 }],
      },
      false,
    )

    expect(result[0].options?.bullet).toBe(true)
  })

  it('adds extra bullet indent when a list item reserves leading badge space', () => {
    const result = toListTextProps({
      text: 'Launch',
      level: 0,
      leadingOffset: 40,
      runs: [{ text: 'Launch', fontSize: 16 }],
    })

    expect(result[0].options?.bullet).toEqual({ indent: 57 })
  })

  it('appends breakLine to last run when more items follow', () => {
    const result = toListTextProps(
      {
        text: 'Line',
        level: 0,
        runs: [
          { text: 'Line ', fontSize: 16 },
          { text: 'Tail', fontSize: 16, bold: true },
        ],
      },
      false,
      true,
    )

    expect(result[0].options?.breakLine).toBeUndefined()
    expect(result[1].options?.breakLine).toBe(true)
  })

  it('omits breakLine on the last item', () => {
    const result = toListTextProps(
      {
        text: 'Last',
        level: 0,
        runs: [{ text: 'Last', fontSize: 16 }],
      },
      false,
      false,
    )

    expect(result[0].options?.breakLine).toBeUndefined()
  })

  it('backgroundColor の run には highlight が設定される — slide 57/59 の strong ハイライト', () => {
    const result = toListTextProps({
      text: 'development efficiency',
      level: 0,
      runs: [
        {
          text: 'Working on ',
          color: 'rgb(0, 0, 0)',
          fontSize: 16,
          fontFamily: 'Arial',
        },
        {
          text: 'development efficiency',
          color: 'rgb(0, 0, 0)',
          fontSize: 16,
          fontFamily: 'Arial',
          backgroundColor: 'rgb(241, 196, 15)', // strong の scoped CSS
        },
        {
          text: ' improvements',
          color: 'rgb(0, 0, 0)',
          fontSize: 16,
          fontFamily: 'Arial',
        },
      ],
    })

    // 通常テキストには highlight なし
    expect(result[0].options?.highlight).toBeUndefined()
    // backgroundColor あり run には highlight が設定される
    expect(result[1].options?.highlight).toBe('F1C40F')
    // 後続テキストにも highlight なし
    expect(result[2].options?.highlight).toBeUndefined()
  })

  it('半透明インラインコード backgroundColor はリスト内でも compositeOverWhite で変換される — slide 21 の <code> ハイライト', () => {
    // Marp デフォルトテーマのインライン <code> は rgba(129,139,152,0.12) を使用する。
    // compositeOverWhite を適用すると rgb(240,241,243) (薄グレー) になり、
    // 全 ch ≤ 248 → highlight = 'f0f1f3'（薄グレーとして表示）となる。
    const result = toListTextProps({
      text: '<',
      level: 0,
      runs: [
        {
          text: '<',
          color: 'rgb(0, 0, 0)',
          fontSize: 16,
          fontFamily: 'Arial',
          backgroundColor: 'rgba(129, 139, 152, 0.12)',
        },
      ],
    })

    expect(result[0].options?.highlight).toBe('F0F1F3')
  })

  it('<br> による継続行は invisible bullet で indent が揃えられる', () => {
    // dom-walker は <li>Line one<br>Line two</li> から
    // [{text:'Line one',...}, {text:'',breakLine:true}, {text:'Line two',...}]
    // を生成する。PptxGenJS の breakLine:true は <a:br/> ではなく新しい <a:p> を
    // 開始するため、継続行が marL=0 に落ちてバレット位置から始まってしまう。
    //
    // 修正方針:
    //   - "Line one" に breakLine:true を付けて arrTexts を空にする
    //   - "Line two" に bullet:{char:'\u200B'} を付けて BulletMarL(342900) を取得
    //   → PptxGenJS が marL=342900 のバレット段落を生成し、テキストが揃う
    const result = toListTextProps({
      text: 'Line one\nLine two',
      level: 0,
      runs: [
        { text: 'Line one', color: 'rgb(0, 0, 0)', fontSize: 16, fontFamily: 'Arial' },
        { text: '', breakLine: true },
        { text: 'Line two', color: 'rgb(0, 0, 0)', fontSize: 16, fontFamily: 'Arial' },
      ],
    })

    // 空の breakLine ランは除去され、2要素になる
    expect(result).toHaveLength(2)
    // 先頭ランに実バレット + indentLevel、かつ breakLine:true で段落を閉じる
    expect(result[0].text).toBe('Line one')
    expect(result[0].options?.bullet).toBe(true)
    expect(result[0].options?.indentLevel).toBe(0)
    expect(result[0].options?.breakLine).toBe(true)
    // 継続ランに invisible bullet + 同じ indentLevel → marL が揃う
    // breakAfter=false なので最後に breakLine は不要
    expect(result[1].text).toBe('Line two')
    expect(result[1].options?.bullet).toEqual({ characterCode: '200B' })
    expect(result[1].options?.indentLevel).toBe(0)
    expect(result[1].options?.breakLine).toBeUndefined()
  })

  it('継続行が複数ある場合もすべて invisible bullet で揃えられる', () => {
    const result = toListTextProps({
      text: 'A\nB\nC',
      level: 1,
      runs: [
        { text: 'A', color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial' },
        { text: '', breakLine: true },
        { text: 'B', color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial' },
        { text: '', breakLine: true },
        { text: 'C', color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial' },
      ],
    })

    expect(result).toHaveLength(3)
    // 各グループの最後の非 lastGroup ランに breakLine が付く
    expect(result[0].options?.bullet).toBe(true)
    expect(result[0].options?.indentLevel).toBe(1)
    expect(result[0].options?.breakLine).toBe(true)
    expect(result[1].options?.bullet).toEqual({ characterCode: '200B' })
    expect(result[1].options?.indentLevel).toBe(1)
    expect(result[1].options?.breakLine).toBe(true)
    // 最終グループ、breakAfter=false → breakLine なし
    expect(result[2].options?.bullet).toEqual({ characterCode: '200B' })
    expect(result[2].options?.indentLevel).toBe(1)
    expect(result[2].options?.breakLine).toBeUndefined()
  })

  it('継続行にも leading badge 用の extra indent を引き継ぐ', () => {
    const result = toListTextProps({
      text: 'Launch\nTomorrow',
      level: 0,
      leadingOffset: 48,
      runs: [
        { text: 'Launch', color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial' },
        { text: '', breakLine: true },
        { text: 'Tomorrow', color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial' },
      ],
    })

    expect(result[0].options?.bullet).toEqual({ indent: 63 })
    expect(result[0].options?.breakLine).toBe(true)
    expect(result[1].options?.bullet).toEqual({ characterCode: '200B', indent: 63 })
  })

  it('同一アイテム内の複数 run すべてに bullet と indentLevel が設定される — PptxGenJS が末尾 <a:pPr> を <a:buNone/> でリセットする問題を防ぐ (slide 52 Item B + emoji)', () => {
    // PptxGenJS v4.x emits <a:pPr> for each TextProp in the same paragraph.
    // LibreOffice uses the *last* <a:pPr>; without propagation the emoji run's
    // pPr resets the bullet with <a:buNone/> causing the bullet to disappear
    // in CI (LibreOffice) even though PowerPoint COM renders it correctly.
    const result = toListTextProps({
      text: 'Item B ✅',
      level: 0,
      runs: [
        {
          text: 'Item B ',
          color: 'rgb(31, 35, 40)',
          fontSize: 29,
          fontFamily: 'Arial',
        },
        {
          text: '✅',
          color: 'rgb(31, 35, 40)',
          fontSize: 29,
          fontFamily: 'Arial',
        },
      ],
    })

    expect(result).toHaveLength(2)
    // Every run in the group must carry bullet and indentLevel
    expect(result[0].options?.bullet).toBe(true)
    expect(result[0].options?.indentLevel).toBe(0)
    expect(result[1].options?.bullet).toBe(true)
    expect(result[1].options?.indentLevel).toBe(0)
    // No spurious breakLine on either run (item is the last item, breakAfter=false)
    expect(result[0].options?.breakLine).toBeUndefined()
    expect(result[1].options?.breakLine).toBeUndefined()
  })

  it('sets numberStartAt on ordered item when startNumber is provided', () => {
    const result = toListTextProps(
      {
        text: 'Fifth',
        level: 0,
        runs: [{ text: 'Fifth', fontSize: 16 }],
      },
      true,   // ordered
      false,  // breakAfter
      undefined,
      undefined,
      5,      // startNumber
    )

    expect(result[0].options?.bullet).toMatchObject({
      type: 'number',
      style: 'arabicPeriod',
      numberStartAt: 5,
    })
  })

  it('omits numberStartAt from ordered item when startNumber is undefined', () => {
    const result = toListTextProps(
      {
        text: 'First',
        level: 0,
        runs: [{ text: 'First', fontSize: 16 }],
      },
      true,   // ordered
    )

    expect((result[0].options?.bullet as any)?.numberStartAt).toBeUndefined()
  })

  it('does not set numberStartAt on unordered item even when startNumber is passed', () => {
    const result = toListTextProps(
      {
        text: 'Bullet',
        level: 0,
        runs: [{ text: 'Bullet', fontSize: 16 }],
      },
      false,  // ordered = false
      false,
      undefined,
      undefined,
      5,      // startNumber ignored for unordered
    )

    expect(result[0].options?.bullet).toBe(true)
  })
})

// -----------------------------------------------------------------------
// placeElement — ordered list startNumber propagation
// -----------------------------------------------------------------------

describe('placeElement — ordered list startNumber', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
    }
  }

  const baseStyle = {
    color: 'rgb(0,0,0)',
    fontSize: 16,
    fontFamily: 'Arial',
    fontWeight: 400,
    textAlign: 'left',
    lineHeight: 0,
  }

  it('passes numberStartAt to first item bullet when startNumber:5 is set', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'list',
      ordered: true,
      startNumber: 5,
      items: [
        { text: 'Item A', level: 0, runs: [{ text: 'Item A', fontSize: 16 }] },
        { text: 'Item B', level: 0, runs: [{ text: 'Item B', fontSize: 16 }] },
      ],
      x: 0,
      y: 0,
      width: 600,
      height: 60,
      style: baseStyle,
    }

    placeElement(mockSlide, el, 1280, 720)

    const textProps: any[] = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    // First item: bullet must include numberStartAt: 5
    expect(textProps[0].options.bullet).toMatchObject({ type: 'number', numberStartAt: 5 })
    // Second item: must include numberStartAt: 6 (PptxGenJS does not auto-increment from startAt)
    expect(textProps[1].options.bullet).toMatchObject({ type: 'number', numberStartAt: 6 })
  })

  it('sets sequential numberStartAt on all items even when startNumber is not set', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'list',
      ordered: true,
      items: [
        { text: 'Item A', level: 0, runs: [{ text: 'Item A', fontSize: 16 }] },
        { text: 'Item B', level: 0, runs: [{ text: 'Item B', fontSize: 16 }] },
      ],
      x: 0,
      y: 0,
      width: 600,
      height: 60,
      style: baseStyle,
    }

    placeElement(mockSlide, el, 1280, 720)

    const textProps: any[] = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    expect((textProps[0].options.bullet as any)?.numberStartAt).toBe(1)
    expect((textProps[1].options.bullet as any)?.numberStartAt).toBe(2)
  })
})

describe('placeElement — image', () => {
  const minimalSlide: SlideData = {
    width: 1280,
    height: 720,
    background: 'rgb(255, 255, 255)',
    backgroundImages: [],
    elements: [],
    notes: '',
  }

  function buildSlideWithImage(src: string) {
    const img: ImageElement = {
      type: 'image',
      src,
      x: 100,
      y: 200,
      width: 400,
      height: 300,
      naturalWidth: 800,
      naturalHeight: 600,
    }
    return buildPptx([{ ...minimalSlide, elements: [img] }])
  }

  it('handles file:// URL without error', () => {
    const fileUrl = 'file:///C:/Users/test/images/photo.png'
    const pptx = buildSlideWithImage(fileUrl)
    expect(pptx).toBeDefined()
  })

  it('handles data: URI without error', () => {
    const dataUri =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
    const pptx = buildSlideWithImage(dataUri)
    expect(pptx).toBeDefined()
  })

  it('handles https URL without error', () => {
    const url = 'https://example.com/image.png'
    const pptx = buildSlideWithImage(url)
    expect(pptx).toBeDefined()
  })
})

describe('buildPptx — background handling', () => {
  const minimalSlide: SlideData = {
    width: 1280,
    height: 720,
    background: 'rgb(255, 255, 255)',
    backgroundImages: [],
    elements: [],
    notes: '',
  }

  it('falls back to white for transparent background', () => {
    const slide: SlideData = {
      ...minimalSlide,
      background: 'rgba(0, 0, 0, 0)',
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })

  it('handles background image without error', () => {
    const slide: SlideData = {
      ...minimalSlide,
      backgroundImages: [
        {
          url: 'https://example.com/bg.png',
          x: 0,
          y: 0,
          width: 1280,
          height: 720,
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })

  it('handles data: URI background image without error', () => {
    const slide: SlideData = {
      ...minimalSlide,
      backgroundImages: [
        {
          url: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUg==',
          x: 0,
          y: 0,
          width: 1280,
          height: 720,
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })

  it('preserves inline code highlight when backgroundSizeContain image is present (slide 76 regression)', () => {
    // Before the fix: a ![bg fit] figure has width = slide width → the
    // old `visualBgMayBeDark` check fired (≥80% width) and suppressed all
    // near-white highlights, making inline <code> elements appear without
    // any background highlight in PPTX.
    //
    // After the fix: backgroundSizeContain images are excluded from the
    // dark-bg check because the actual image is letterboxed in the center
    // and the margins remain on the CSS background (white).
    // → visualBgMayBeDark = false → highlight preserved.
    //
    // We verify this via toTextProps directly, since the new logic determines
    // visualBgMayBeDark before calling placeElement which calls toTextProps.
    // With visualBgMayBeDark=false and dark text, the highlight must be defined.
    const result = toTextProps(
      {
        text: '![bg fit]',
        color: 'rgb(51, 51, 51)', // dark text on white margin area
        fontSize: 16,
        backgroundColor: 'rgba(129, 139, 152, 0.12)', // Marp inline <code> bg
      },
      'rgb(255, 255, 255)',
      false, // backgroundSizeContain → not treated as dark bg → visualBgMayBeDark=false
    )
    expect(result.options?.highlight).toBeDefined()
    expect(result.options?.highlight).toBe('F0F1F3')
  })

  it('shows inline code highlight even when element overlaps a split background image (slide 75 structural fidelity)', () => {
    // Design principle: structural fidelity > visual approximation.
    // When a split ![bg] places text over a colored image region, compositing
    // the rgba(0,0,0,0.12) code background over white gives a near-white box
    // (#F0F1F3) that appears as a white rectangle on the image.  The correct
    // response is to accept this visual imperfection rather than suppress the
    // highlight — the HTML says "<code>", so PPTX should say "code highlight".
    //
    // Per-element overlap suppression (removed in the design revision after
    // ADR-34) was an ad-hoc case that violated the browser-is-source-of-truth
    // principle.  This test ensures it is NOT reintroduced.
    //
    // We verify via toTextProps with visualBgMayBeDark=false (split images are
    // each ~50% wide, below the ≥80% full-slide threshold).
    const result = toTextProps(
      {
        text: '![bg]',
        color: 'rgb(51, 51, 51)', // dark text over colored split bg image
        fontSize: 16,
        backgroundColor: 'rgba(129, 139, 152, 0.12)', // Marp inline <code> bg
      },
      'rgb(255, 255, 255)',
      false, // split images are < 80% wide → full-slide dark check does not fire
    )
    // Highlight must be present — suppressing it would violate structural fidelity
    expect(result.options?.highlight).toBeDefined()
    expect(result.options?.highlight).toBe('F0F1F3')
  })
})

describe('placeElement — table with transparent cells', () => {
  const minimalSlide: SlideData = {
    width: 1280,
    height: 720,
    background: 'rgb(255, 255, 255)',
    backgroundImages: [],
    elements: [],
    notes: '',
  }

  it('handles transparent cell background without error', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'table',
          rows: [
            {
              cells: [
                {
                  text: 'Header',
                  runs: [
                    {
                      text: 'Header',
                      color: 'rgb(0, 0, 0)',
                      fontSize: 16,
                      bold: true,
                    },
                  ],
                  isHeader: true,
                  style: {
                    color: 'rgb(0, 0, 0)',
                    backgroundColor: 'rgba(0, 0, 0, 0)',
                    fontSize: 16,
                    fontFamily: 'Arial',
                    fontWeight: 700,
                    textAlign: 'left',
                    borderColor: 'rgb(200, 200, 200)',
                  },
                },
              ],
            },
            {
              cells: [
                {
                  text: 'Cell',
                  runs: [{ text: 'Cell', color: 'rgb(0, 0, 0)', fontSize: 16 }],
                  isHeader: false,
                  style: {
                    color: 'rgb(0, 0, 0)',
                    backgroundColor: 'rgba(0, 0, 0, 0)',
                    fontSize: 16,
                    fontFamily: 'Arial',
                    fontWeight: 400,
                    textAlign: 'left',
                    borderColor: 'rgb(200, 200, 200)',
                  },
                },
              ],
            },
          ],
          x: 70,
          y: 200,
          width: 600,
          height: 80,
          style: {
            color: 'rgb(0, 0, 0)',
            fontSize: 16,
            fontFamily: 'Arial',
            fontWeight: 400,
            textAlign: 'left',
            lineHeight: 24,
          },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })

  it('bolds cells with fontWeight >= 700', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'table',
          rows: [
            {
              cells: [
                {
                  text: 'Bold cell',
                  runs: [],
                  isHeader: false,
                  style: {
                    color: 'rgb(0, 0, 0)',
                    backgroundColor: 'rgb(240, 240, 240)',
                    fontSize: 16,
                    fontFamily: 'Arial',
                    fontWeight: 700,
                    textAlign: 'left',
                    borderColor: 'rgb(200, 200, 200)',
                  },
                },
              ],
            },
          ],
          x: 70,
          y: 200,
          width: 600,
          height: 40,
          style: {
            color: 'rgb(0, 0, 0)',
            fontSize: 16,
            fontFamily: 'Arial',
            fontWeight: 400,
            textAlign: 'left',
            lineHeight: 24,
          },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })
})

describe('placeElement — blockquote with border', () => {
  const minimalSlide: SlideData = {
    width: 1280,
    height: 720,
    background: 'rgb(255, 255, 255)',
    backgroundImages: [],
    elements: [],
    notes: '',
  }

  it('places blockquote with left border without error', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'blockquote',
          runs: [{ text: 'Quote text', color: 'rgb(0, 0, 0)', fontSize: 16 }],
          x: 70,
          y: 100,
          width: 600,
          height: 40,
          style: {
            color: 'rgb(0, 0, 0)',
            fontSize: 16,
            fontFamily: 'Arial',
            fontWeight: 400,
            textAlign: 'left',
            lineHeight: 24,
          },
          borderLeft: { width: 4, color: 'rgb(100, 100, 100)' },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })
})

describe('placeElement — code with syntax runs', () => {
  const minimalSlide: SlideData = {
    width: 1280,
    height: 720,
    background: 'rgb(255, 255, 255)',
    backgroundImages: [],
    elements: [],
    notes: '',
  }

  it('places syntax-highlighted code block without error', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'code',
          text: 'const x = 1;',
          language: 'javascript',
          runs: [
            {
              text: 'const',
              color: 'rgb(198, 120, 221)',
              fontSize: 14,
              bold: true,
            },
            { text: ' x = ', color: 'rgb(200, 200, 200)', fontSize: 14 },
            { text: '1', color: 'rgb(209, 154, 102)', fontSize: 14 },
            { text: ';', color: 'rgb(200, 200, 200)', fontSize: 14 },
          ],
          x: 70,
          y: 200,
          width: 600,
          height: 80,
          style: {
            color: 'rgb(200, 200, 200)',
            fontSize: 14,
            fontFamily: 'monospace',
            fontWeight: 400,
            textAlign: 'left',
            lineHeight: 20,
            backgroundColor: 'rgb(40, 44, 52)',
          },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })
})

describe('placeElement — heading with border', () => {
  const minimalSlide: SlideData = {
    width: 1280,
    height: 720,
    background: 'rgb(255, 255, 255)',
    backgroundImages: [],
    elements: [],
    notes: '',
  }
  const baseStyle = {
    color: 'rgb(44, 62, 80)',
    fontSize: 40,
    fontFamily: 'Arial',
    fontWeight: 700,
    textAlign: 'left' as const,
    lineHeight: 48,
  }

  it('places heading with border-bottom without error', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'heading',
          level: 1,
          runs: [{ text: 'Title', color: 'rgb(44, 62, 80)', fontSize: 40 }],
          x: 70,
          y: 80,
          width: 1140,
          height: 57,
          style: baseStyle,
          borderBottom: { width: 2, color: 'rgb(39, 174, 96)' },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })

  it('places heading with border-left without errors', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'heading',
          level: 2,
          runs: [{ text: 'Section', color: 'rgb(39, 174, 96)', fontSize: 30 }],
          x: 70,
          y: 52,
          width: 1140,
          height: 36,
          style: { ...baseStyle, fontSize: 30, lineHeight: 36 },
          borderLeft: { width: 4, color: 'rgb(39, 174, 96)' },
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })

  it('places heading without border without errors', () => {
    const slide: SlideData = {
      ...minimalSlide,
      elements: [
        {
          type: 'heading',
          level: 1,
          runs: [{ text: 'No border', color: 'rgb(0,0,0)', fontSize: 40 }],
          x: 70,
          y: 80,
          width: 1140,
          height: 57,
          style: baseStyle,
        },
      ],
    }
    const pptx = buildPptx([slide])
    expect(pptx).toBeDefined()
  })
})

// ---------------------------------------------------------------------------
// placeElement — heading border-left text shift
// ---------------------------------------------------------------------------

describe('placeElement — heading border-left text offset', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  const baseStyle = {
    color: 'rgb(41, 128, 185)',
    fontSize: 30,
    fontFamily: 'Arial',
    fontWeight: 700,
    textAlign: 'left' as const,
    lineHeight: 36,
  }

  it('h2 border-left: text box shifts right by border width', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'heading',
      level: 2,
      runs: [
        { text: 'Section heading', color: 'rgb(41,128,185)', fontSize: 30 },
      ],
      x: 70,
      y: 52,
      width: 1140,
      height: 36,
      style: baseStyle,
      borderLeft: { width: 4, color: 'rgb(41, 128, 185)' },
    }
    placeElement(mockSlide, el, 1280, 720)

    const textCall = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    const bwIn = 4 / 96 // 4px → inches
    // text x shifted right by border width (x + bw)
    expect(textCall.x).toBeCloseTo(70 / 96 + bwIn, 6)
    // Full-width heading: width extends to slide boundary (slideW - x - 16px buffer) minus border
    expect(textCall.w).toBeCloseTo((1280 - 70 - 16 - 4) / 96, 6)
  })

  it('h2 border-left: border rect drawn before text (z-order)', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'heading',
      level: 2,
      runs: [
        { text: 'Section heading', color: 'rgb(41,128,185)', fontSize: 30 },
      ],
      x: 70,
      y: 52,
      width: 1140,
      height: 36,
      style: baseStyle,
      borderLeft: { width: 4, color: 'rgb(41, 128, 185)' },
    }
    placeElement(mockSlide, el, 1280, 720)

    const addShapeOrder = (mockSlide.addShape as jest.Mock).mock
      .invocationCallOrder[0]
    const addTextOrder = (mockSlide.addText as jest.Mock).mock
      .invocationCallOrder[0]
    // shape (border bar) must be drawn before text so text renders on top
    expect(addShapeOrder).toBeLessThan(addTextOrder)
  })

  it('h2 without border-left: text box stays at original x', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'heading',
      level: 2,
      runs: [{ text: 'No decoration', color: 'rgb(0,0,0)', fontSize: 30 }],
      x: 70,
      y: 52,
      width: 1140,
      height: 36,
      style: baseStyle,
      // no borderLeft
    }
    placeElement(mockSlide, el, 1280, 720)

    const textCall = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(textCall.x).toBeCloseTo(70 / 96, 6)
    // Full-width heading: width extends to slide boundary (slideW - x - 16px buffer)
    expect(textCall.w).toBeCloseTo((1280 - 70 - 16) / 96, 6)
  })
})

// ---------------------------------------------------------------------------
// placeElement — text height clamping
// ---------------------------------------------------------------------------

describe('placeElement — text height clamping', () => {
  // Slide is 720px tall. A paragraph at y=680 with height=80 would extend
  // 40px below the slide boundary. placeElement() should clamp text-type
  // elements so y + h ≤ slideH. Images are intentionally excluded.
  // pxToInches converts at 96 dpi, so 1px = 1/96 in.

  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  it('clamps paragraph height when it would overflow the slide bottom', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'paragraph',
      runs: [
        {
          text: 'Overflowing text',
          color: 'rgb(0,0,0)',
          fontSize: 16,
          fontFamily: 'Arial',
          bold: false,
        },
      ],
      x: 0,
      y: 680, // near bottom; 680 + 80 = 760 > 720
      width: 1280,
      height: 80,
      style: {
        textAlign: 'left',
        fontFamily: 'Arial',
        fontSize: 16,
        fontWeight: 400,
        color: 'rgb(0,0,0)',
        lineHeight: 24,
      },
    }

    placeElement(mockSlide, el, 1280, 720)

    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    // Clamped h = (720 - 680) / 96 ≈ 0.4167 in — not the raw 80/96 ≈ 0.8333
    expect(opts.h).toBeCloseTo(40 / 96, 4)
    expect(opts.h).toBeLessThan(80 / 96)
  })

  it('does not clamp image height even when it overflows slide bounds', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'image',
      src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      x: 0,
      y: 680,
      width: 1280,
      height: 80,
    }

    placeElement(mockSlide, el, 1280, 720)

    const opts = (mockSlide.addImage as jest.Mock).mock.calls[0][0]
    // Image h must NOT be clamped
    expect(opts.h).toBeCloseTo(80 / 96, 4)
  })
})

// -----------------------------------------------------------------------
// placeElement — lineSpacingMultiple from CSS line-height
// -----------------------------------------------------------------------

describe('placeElement — lineSpacingMultiple from CSS line-height', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  const baseStyle = {
    textAlign: 'left' as const,
    fontFamily: 'Arial',
    fontWeight: 400,
    color: 'rgb(0,0,0)',
  }

  it('applies lineSpacingMultiple = lineHeight/fontSize to paragraph', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'paragraph',
      // CJK text intentionally used to test font rendering path
      runs: [
        {
          text: 'Test',
          color: 'rgb(0,0,0)',
          fontSize: 16,
          fontFamily: 'Arial',
          bold: false,
        },
      ],
      x: 0,
      y: 0,
      width: 600,
      height: 40,
      style: { ...baseStyle, fontSize: 16, lineHeight: 24 }, // 24/16 = 1.5; h/lh=40/24=1.67>1.5 → multiLine → /1.20 = 1.25
    }
    placeElement(mockSlide, el, 1280, 720)
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(opts.lineSpacingMultiple).toBeCloseTo(1.25, 2)
  })

  it('applies lineSpacingMultiple to heading', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'heading',
      level: 2,
      runs: [
        {
          text: 'Heading',
          color: 'rgb(0,0,0)',
          fontSize: 32,
          fontFamily: 'Arial',
          bold: true,
        },
      ],
      x: 0,
      y: 0,
      width: 600,
      height: 60,
      style: { ...baseStyle, fontSize: 32, lineHeight: 40 }, // 40/32 = 1.25 → /1.15 = 1.09
    }
    placeElement(mockSlide, el, 1280, 720)
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(opts.lineSpacingMultiple).toBeCloseTo(1.09, 2)
  })

  it('applies lineSpacingMultiple to list', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'list',
      ordered: false,
      items: [
        {
          text: 'item',
          level: 0,
          runs: [
            {
              text: 'item',
              color: 'rgb(0,0,0)',
              fontSize: 16,
              fontFamily: 'Arial',
              bold: false,
            },
          ],
        },
      ],
      x: 0,
      y: 0,
      width: 600,
      height: 40,
      style: { ...baseStyle, fontSize: 16, lineHeight: 22 }, // 22/16 = 1.375; h/lh=40/22=1.82>1.5 → multiLine → /1.20 = 1.15
    }
    placeElement(mockSlide, el, 1280, 720)
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(opts.lineSpacingMultiple).toBeCloseTo(1.15, 2)
  })

  it('omits lineSpacingMultiple when lineHeight is 0 (normal)', () => {
    const mockSlide = makeMockSlide() as any
    const el: any = {
      type: 'paragraph',
      runs: [
        {
          text: 'Test',
          color: 'rgb(0,0,0)',
          fontSize: 16,
          fontFamily: 'Arial',
          bold: false,
        },
      ],
      x: 0,
      y: 0,
      width: 600,
      height: 40,
      style: { ...baseStyle, fontSize: 16, lineHeight: 0 }, // lineHeight=0 → undefined
    }
    placeElement(mockSlide, el, 1280, 720)
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(opts.lineSpacingMultiple).toBeUndefined()
  })
})

// -----------------------------------------------------------------------
// placeElement — container strips matching highlight from children
// -----------------------------------------------------------------------

describe('placeElement — container child highlight strip', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  it('strips text highlight that matches container background', () => {
    const mockSlide = makeMockSlide()
    const childParagraph: any = {
      type: 'paragraph',
      runs: [
        {
          text: 'same color',
          color: 'rgb(0,0,0)',
          fontSize: 16,
          backgroundColor: 'rgb(52,152,219)',
        },
        {
          text: 'different',
          color: 'rgb(0,0,0)',
          fontSize: 16,
          backgroundColor: 'rgb(241,196,15)',
        },
      ],
      x: 80,
      y: 100,
      width: 400,
      height: 30,
      style: { textAlign: 'left', fontSize: 16, lineHeight: 0 },
    }
    const el: any = {
      type: 'container',
      children: [childParagraph],
      x: 70,
      y: 90,
      width: 500,
      height: 200,
      style: { backgroundColor: 'rgb(52, 152, 219)' },
    }
    placeElement(mockSlide, el, 1280, 720)

    // The child paragraph's first run should have had its backgroundColor stripped
    expect(childParagraph.runs[0].backgroundColor).toBeUndefined()
    // The second run with a different color should be preserved
    expect(childParagraph.runs[1].backgroundColor).toBe('rgb(241,196,15)')
  })

  it('preserves highlight when container has no background', () => {
    const mockSlide = makeMockSlide()
    const childParagraph: any = {
      type: 'paragraph',
      runs: [
        {
          text: 'highlighted',
          color: 'rgb(0,0,0)',
          fontSize: 16,
          backgroundColor: 'rgb(241,196,15)',
        },
      ],
      x: 80,
      y: 100,
      width: 400,
      height: 30,
      style: { textAlign: 'left', fontSize: 16, lineHeight: 0 },
    }
    const el: any = {
      type: 'container',
      children: [childParagraph],
      x: 70,
      y: 90,
      width: 500,
      height: 200,
      style: { backgroundColor: 'transparent' },
    }
    placeElement(mockSlide, el, 1280, 720)

    expect(childParagraph.runs[0].backgroundColor).toBe('rgb(241,196,15)')
  })
})

// ---------------------------------------------------------------------------
// placeElement — paragraph text inset (margin) for asymmetric padding
//
// PptxGenJS maps margin[0]→lIns, [1]→rIns, [2]→bIns, [3]→tIns.
// Our computeTextInset must return [left, right, bottom, top] so that the
// OOXML tIns / lIns values match the CSS paddingTop / paddingLeft.
// ---------------------------------------------------------------------------

describe('placeElement — paragraph text inset is correct for asymmetric padding', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  it('asymmetric padding: margin[3] (→tIns) = paddingTop * 0.75pt, margin[0] (→lIns) = paddingLeft * 0.75pt', () => {
    const mockSlide = makeMockSlide()
    // padding: 10px top/bottom, 24px left/right (like "Input data" button)
    const el: any = {
      type: 'paragraph',
      runs: [{ text: 'Input data', color: 'rgb(255,255,255)', fontSize: 16 }],
      x: 545,
      y: 318,
      width: 190,
      height: 64,
      style: {
        color: 'rgb(255,255,255)',
        fontSize: 16,
        fontFamily: 'Arial',
        fontWeight: 400,
        textAlign: 'center' as const,
        lineHeight: 24,
        paddingTop: 10,
        paddingRight: 24,
        paddingBottom: 10,
        paddingLeft: 24,
      },
      valign: 'top' as const,
    }
    placeElement(mockSlide, el, 1280, 720)

    const textOpts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    const margin = textOpts.margin as [number, number, number, number]
    // PptxGenJS order: [0]=lIns, [1]=rIns, [2]=bIns, [3]=tIns
    expect(margin[3]).toBeCloseTo(10 * 0.75, 4) // tIns = paddingTop * 0.75pt
    expect(margin[0]).toBeCloseTo(24 * 0.75, 4) // lIns = paddingLeft * 0.75pt
    expect(margin[1]).toBeCloseTo(24 * 0.75, 4) // rIns = paddingRight * 0.75pt
    expect(margin[2]).toBeCloseTo(10 * 0.75, 4) // bIns = paddingBottom * 0.75pt
  })

  it('symmetric padding: margin values all equal regardless of order', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'paragraph',
      runs: [{ text: 'Hello', color: 'rgb(0,0,0)', fontSize: 16 }],
      x: 0,
      y: 0,
      width: 200,
      height: 50,
      style: {
        color: 'rgb(0,0,0)',
        fontSize: 16,
        fontFamily: 'Arial',
        fontWeight: 400,
        textAlign: 'left' as const,
        lineHeight: 24,
        paddingTop: 12,
        paddingRight: 12,
        paddingBottom: 12,
        paddingLeft: 12,
      },
    }
    placeElement(mockSlide, el, 1280, 720)

    const textOpts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    const margin = textOpts.margin as [number, number, number, number]
    const expected = 12 * 0.75
    expect(margin[0]).toBeCloseTo(expected, 4)
    expect(margin[1]).toBeCloseTo(expected, 4)
    expect(margin[2]).toBeCloseTo(expected, 4)
    expect(margin[3]).toBeCloseTo(expected, 4)
  })
})

// ---------------------------------------------------------------------------
// placeElement — paragraph width extension heuristic
// ---------------------------------------------------------------------------

describe('placeElement — paragraph width extension (DirectWrite compensation)', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  const baseStyle = {
    color: 'rgb(0,0,0)',
    fontSize: 15,
    fontFamily: 'Arial',
    fontWeight: 400,
    textAlign: 'left' as const,
    lineHeight: 22,
  }

  it('extends paragraph width by 5% (DIRECTWRITE_COL_WIDTH_FACTOR)', () => {
    // x=79, width=898: extended = 898 * 1.05 = 942.9, cap = 1280-79-4 = 1197
    // Expected: 942.9 / 96
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'paragraph',
      runs: [{ text: 'Long chat bubble text', fontSize: 15 }],
      x: 79,
      y: 200,
      width: 898,
      height: 30,
      style: baseStyle,
    }
    placeElement(mockSlide, el, 1280, 720)

    const w = (mockSlide.addText as jest.Mock).mock.calls[0][1].w as number
    const expectedW = Math.min(898 * 1.05, 1280 - 79 - 4) / 96
    expect(w).toBeCloseTo(expectedW, 5)
  })

  it('extends narrow paragraph by 5% as well', () => {
    // x=79, width=400: extended = 400 * 1.05 = 420, cap = 1280-79-4 = 1197
    // Expected: 420 / 96
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'paragraph',
      runs: [{ text: 'Short paragraph', fontSize: 15 }],
      x: 79,
      y: 100,
      width: 400,
      height: 24,
      style: baseStyle,
    }
    placeElement(mockSlide, el, 1280, 720)

    const w = (mockSlide.addText as jest.Mock).mock.calls[0][1].w as number
    expect(w).toBeCloseTo(400 * 1.05 / 96, 5)
  })

  it('extends short far-right paragraph by 5%', () => {
    // x=1000, width=200: extended = 200 * 1.05 = 210, cap = 1280-1000-4 = 276
    // Expected: 210 / 96
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'paragraph',
      runs: [{ text: 'Tiny', fontSize: 15 }],
      x: 1000,
      y: 100,
      width: 200,
      height: 24,
      style: baseStyle,
    }
    placeElement(mockSlide, el, 1280, 720)

    const w = (mockSlide.addText as jest.Mock).mock.calls[0][1].w as number
    expect(w).toBeCloseTo(210 / 96, 5)
  })

  it('caps extension at slideW − x − 4 to avoid slide overflow', () => {
    // x=79, width=1185: extended = 1185 * 1.05 = 1244.25, cap = 1280-79-4 = 1197
    // Expected: 1197 / 96
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'paragraph',
      runs: [{ text: 'Nearly full width paragraph', fontSize: 15 }],
      x: 79,
      y: 100,
      width: 1185,
      height: 24,
      style: baseStyle,
    }
    placeElement(mockSlide, el, 1280, 720)

    const w = (mockSlide.addText as jest.Mock).mock.calls[0][1].w as number
    const expectedW = Math.min(1185 * 1.05, 1280 - 79 - 4) / 96
    expect(w).toBeCloseTo(expectedW, 5)
  })
})

// ---------------------------------------------------------------------------
// T1: container borderStyle → PptxGenJS dashType mapping
// ---------------------------------------------------------------------------

describe('placeElement — container borderDashType mapping', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  it('dashed borderStyle produces dashType:dash on addShape line', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'container',
      children: [],
      x: 50, y: 50, width: 400, height: 100,
      style: {
        backgroundColor: 'rgb(255,244,232)',
        borderWidth: 2,
        borderColor: 'rgb(200,0,0)',
        borderStyle: 'dashed',
      },
    }
    placeElement(mockSlide, el, 1280, 720)
    const shapeCall = (mockSlide.addShape as jest.Mock).mock.calls[0]
    expect(shapeCall).toBeDefined()
    const opts = shapeCall[1]
    expect(opts.line?.dashType).toBe('dash')
  })

  it('dotted borderStyle produces dashType:sysDot on addShape line', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'container',
      children: [],
      x: 50, y: 50, width: 400, height: 100,
      style: {
        backgroundColor: 'rgb(255,244,232)',
        borderWidth: 1,
        borderColor: 'rgb(100,100,100)',
        borderStyle: 'dotted',
      },
    }
    placeElement(mockSlide, el, 1280, 720)
    const shapeCall = (mockSlide.addShape as jest.Mock).mock.calls[0]
    expect(shapeCall).toBeDefined()
    const opts = shapeCall[1]
    expect(opts.line?.dashType).toBe('sysDot')
  })

  it('solid borderStyle does NOT produce dashType', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'container',
      children: [],
      x: 50, y: 50, width: 400, height: 100,
      style: {
        backgroundColor: 'rgb(255,244,232)',
        borderWidth: 2,
        borderColor: 'rgb(0,0,0)',
        borderStyle: 'solid',
      },
    }
    placeElement(mockSlide, el, 1280, 720)
    const shapeCall = (mockSlide.addShape as jest.Mock).mock.calls[0]
    expect(shapeCall).toBeDefined()
    const opts = shapeCall[1]
    expect(opts.line?.dashType).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// T2: heading paddingLeft → margin inset
// ---------------------------------------------------------------------------

describe('placeElement — heading padding produces text inset', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  it('heading with paddingLeft produces non-zero margin[0] (lIns)', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'heading',
      level: 2,
      runs: [{ text: 'Heading', color: 'rgb(0,0,0)', fontSize: 28 }],
      x: 30, y: 50, width: 600, height: 40,
      style: {
        color: 'rgb(0,0,0)', fontSize: 28, fontFamily: 'Arial',
        fontWeight: 700, textAlign: 'left', lineHeight: 34,
        paddingTop: 8, paddingRight: 0, paddingBottom: 8, paddingLeft: 16,
      },
    }
    placeElement(mockSlide, el, 1280, 720)
    const textCall = (mockSlide.addText as jest.Mock).mock.calls[0]
    const opts = textCall[1]
    // margin = [lIns, rIns, tIns, bIns] where lIns = paddingLeft * 0.75pt
    expect(opts.margin).toBeDefined()
    expect(opts.margin[0]).toBeGreaterThan(0) // lIns from paddingLeft
    expect(opts.margin[2]).toBeGreaterThan(0) // tIns from paddingTop
  })

  it('heading without padding produces margin 0', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'heading',
      level: 1,
      runs: [{ text: 'Title', color: 'rgb(0,0,0)', fontSize: 40 }],
      x: 0, y: 0, width: 600, height: 50,
      style: {
        color: 'rgb(0,0,0)', fontSize: 40, fontFamily: 'Arial',
        fontWeight: 700, textAlign: 'left', lineHeight: 48,
      },
    }
    placeElement(mockSlide, el, 1280, 720)
    const textCall = (mockSlide.addText as jest.Mock).mock.calls[0]
    const opts = textCall[1]
    expect(opts.margin).toBe(0)
  })
})

// ---------------------------------------------------------------------------
// T7: table cell margin reduction
// ---------------------------------------------------------------------------

describe('placeElement — table cell margin', () => {
  function makeMockSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  it('table placement sets asymmetric margin [0.1, 0.05, 0.1, 0.05] (top/bottom larger than left/right)', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'table',
      rows: [
        {
          cells: [
            { text: 'A', runs: [{ text: 'A', fontSize: 14 }], style: { fontWeight: 400, color: 'rgb(0,0,0)', backgroundColor: 'rgb(255,255,255)' } },
            { text: 'B', runs: [{ text: 'B', fontSize: 14 }], style: { fontWeight: 400, color: 'rgb(0,0,0)', backgroundColor: 'rgb(255,255,255)' } },
          ],
        },
      ],
      x: 50, y: 100, width: 600, height: 40,
      style: { backgroundColor: 'transparent' },
    }
    placeElement(mockSlide, el, 1280, 720)
    const tableCall = (mockSlide.addTable as jest.Mock).mock.calls[0]
    expect(tableCall).toBeDefined()
    const opts = tableCall[1]
    // top=0.1, right=0.05, bottom=0.1, left=0.05 (CSS order)
    // - top/bottom 0.1in ≈ 9.6px: improves row height vs browser 6px padding
    // - left/right 0.05in ≈ 4.8px: keeps text area wide to prevent header wrapping
    expect(opts.margin).toEqual([0.1, 0.05, 0.1, 0.05])
  })

  it('table cell run with backgroundColor emits highlight property', () => {
    const mockSlide = makeMockSlide()
    const el: any = {
      type: 'table',
      rows: [
        {
          cells: [
            {
              text: 'Code cell',
              runs: [
                { text: 'Code ', color: 'rgb(0,0,0)', fontSize: 16, backgroundColor: 'rgba(127,139,152,0.12)' },
                { text: 'cell', color: 'rgb(0,0,0)', fontSize: 16 },
              ],
              style: { fontWeight: 400, color: 'rgb(0,0,0)', backgroundColor: 'rgb(255,255,255)', fontSize: 16, fontFamily: 'Arial', textAlign: 'center' },
            },
          ],
        },
      ],
      x: 50, y: 100, width: 600, height: 40,
      style: { backgroundColor: 'transparent' },
    }
    placeElement(mockSlide, el, 1280, 720)
    const tableCall = (mockSlide.addTable as jest.Mock).mock.calls[0]
    expect(tableCall).toBeDefined()
    const firstRunOpts = tableCall[0][0][0].text[0].options
    // rgba(127,139,152,0.12) over white → composited ≈ rgb(242,243,244), delta>10 → highlight shown
    expect(firstRunOpts.highlight).toBeDefined()
    // Second run has no backgroundColor → no highlight
    const secondRunOpts = tableCall[0][0][0].text[1].options
    expect(secondRunOpts.highlight).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// visualBgMayBeDark: CSS gradient placeholder must NOT suppress code highlights
// placeElement を直接使い、visualBgMayBeDark フラグの有無で highlight 有無を確認する
// ---------------------------------------------------------------------------

describe('placeElement — visualBgMayBeDark で inline code highlight の有無が変わる', () => {
  function makeSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  const inlineCodeParagraph: any = {
    type: 'paragraph',
    runs: [
      { text: 'before ', color: 'rgb(0,0,0)', fontSize: 16 },
      // inline code: rgba over white background → 合成後は約 rgb(236,237,238)
      { text: 'code', color: 'rgb(0,0,0)', fontSize: 16, backgroundColor: 'rgba(129,139,152,0.12)' },
    ],
    x: 70, y: 200, width: 1000, height: 40,
    style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
  }

  it('visualBgMayBeDark=false（CSS gradient）のとき inline code highlight が出力される', () => {
    const mockSlide = makeSlide()
    placeElement(mockSlide, inlineCodeParagraph, 1280, 720, 'rgb(255,255,255)', false)

    const calls = (mockSlide.addText as jest.Mock).mock.calls
    const textArr = calls.find((c) => Array.isArray(c[0]) && c[0].some((t: any) => t.text === 'code'))
    expect(textArr).toBeDefined()
    const codeRun = textArr![0].find((t: any) => t.text === 'code')
    expect(codeRun.options.highlight).toBeDefined()
  })

  it('visualBgMayBeDark=true（実画像背景）のとき明るい inline code highlight は抑制される', () => {
    const mockSlide = makeSlide()
    placeElement(mockSlide, inlineCodeParagraph, 1280, 720, 'rgb(255,255,255)', true)

    const calls = (mockSlide.addText as jest.Mock).mock.calls
    const textArr = calls.find((c) => Array.isArray(c[0]) && c[0].some((t: any) => t.text === 'code'))
    expect(textArr).toBeDefined()
    const codeRun = textArr![0].find((t: any) => t.text === 'code')
    expect(codeRun.options.highlight).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// placeElement — table colspan and rowspan
// ---------------------------------------------------------------------------

describe('placeElement — table colspan and rowspan', () => {
  const baseStyle = {
    color: 'rgb(0,0,0)',
    backgroundColor: 'rgb(240,240,240)',
    fontSize: 16,
    fontFamily: 'Arial',
    fontWeight: 400,
    textAlign: 'left',
    borderColor: 'rgb(200,200,200)',
  }

  function makeSlide() {
    return {
      addText: jest.fn(),
      addShape: jest.fn(),
      addImage: jest.fn(),
      addTable: jest.fn(),
      addNotes: jest.fn(),
    } as unknown as any
  }

  it('passes colspan:2 to addTable cell options for a merged header cell (runs branch)', () => {
    const el: any = {
      type: 'table',
      x: 70, y: 100, width: 600, height: 60,
      rows: [
        {
          cells: [
            {
              text: 'Merged Header',
              runs: [{ text: 'Merged Header', color: 'rgb(0,0,0)', fontSize: 16 }],
              isHeader: true,
              colspan: 2,
              style: baseStyle,
            },
          ],
        },
      ],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide, el, 1280, 720, 'rgb(255,255,255)', false)

    const addTableCalls = (mockSlide.addTable as jest.Mock).mock.calls
    expect(addTableCalls).toHaveLength(1)
    const rows = addTableCalls[0][0] as any[][]
    expect(rows[0][0].options.colspan).toBe(2)
  })

  it('passes rowspan:2 to addTable cell options for a row-spanning cell (runs branch)', () => {
    const el: any = {
      type: 'table',
      x: 70, y: 100, width: 600, height: 80,
      rows: [
        {
          cells: [
            {
              text: 'Row Header',
              runs: [{ text: 'Row Header', color: 'rgb(0,0,0)', fontSize: 16 }],
              isHeader: false,
              rowspan: 2,
              style: baseStyle,
            },
            {
              text: 'Top Right',
              runs: [{ text: 'Top Right', color: 'rgb(0,0,0)', fontSize: 16 }],
              isHeader: false,
              style: baseStyle,
            },
          ],
        },
        {
          cells: [
            {
              text: 'Bottom Right',
              runs: [{ text: 'Bottom Right', color: 'rgb(0,0,0)', fontSize: 16 }],
              isHeader: false,
              style: baseStyle,
            },
          ],
        },
      ],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide, el, 1280, 720, 'rgb(255,255,255)', false)

    const addTableCalls = (mockSlide.addTable as jest.Mock).mock.calls
    const rows = addTableCalls[0][0] as any[][]
    expect(rows[0][0].options.rowspan).toBe(2)
  })

  it('omits colspan from options when value is 1 (runs branch)', () => {
    const el: any = {
      type: 'table',
      x: 70, y: 100, width: 600, height: 40,
      rows: [
        {
          cells: [
            {
              text: 'Normal',
              runs: [{ text: 'Normal', color: 'rgb(0,0,0)', fontSize: 16 }],
              isHeader: false,
              colspan: 1,
              style: baseStyle,
            },
          ],
        },
      ],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide, el, 1280, 720, 'rgb(255,255,255)', false)

    const rows = (mockSlide.addTable as jest.Mock).mock.calls[0][0] as any[][]
    expect(rows[0][0].options.colspan).toBeUndefined()
  })

  it('passes colspan:2 to addTable cell options for a merged cell (plain-text fallback branch)', () => {
    const el: any = {
      type: 'table',
      x: 70, y: 100, width: 600, height: 40,
      rows: [
        {
          cells: [
            {
              text: 'Merged',
              runs: [],  // empty runs → fallback branch
              isHeader: false,
              colspan: 2,
              style: baseStyle,
            },
          ],
        },
      ],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide, el, 1280, 720, 'rgb(255,255,255)', false)

    const rows = (mockSlide.addTable as jest.Mock).mock.calls[0][0] as any[][]
    expect(rows[0][0].options.colspan).toBe(2)
  })

  it('preserves breakLine runs inside table cells for multi-line content', () => {
    const el: any = {
      type: 'table',
      x: 70, y: 100, width: 600, height: 80,
      rows: [
        {
          cells: [
            {
              text: 'Line1\nLine2',
              runs: [
                { text: 'Line1', color: 'rgb(0,0,0)', fontSize: 16 },
                { text: '', breakLine: true },
                { text: 'Line2', color: 'rgb(0,0,0)', fontSize: 16 },
              ],
              isHeader: false,
              style: baseStyle,
            },
          ],
        },
      ],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide, el, 1280, 720, 'rgb(255,255,255)', false)

    const rows = (mockSlide.addTable as jest.Mock).mock.calls[0][0] as any[][]
    const cellTextArray = rows[0][0].text
    expect(cellTextArray).toHaveLength(3)
    expect(cellTextArray[0].text).toBe('Line1')
    expect(cellTextArray[1]).toEqual({ text: '', options: { breakLine: true } })
    expect(cellTextArray[2].text).toBe('Line2')
  })
})

// ── Text grouping ─────────────────────────────────────────────────────
describe('groupAdjacentTextElements', () => {
  const baseStyle = {
    color: 'rgb(0,0,0)',
    fontSize: 24,
    fontFamily: 'Arial',
    fontWeight: 400,
    textAlign: 'left' as const,
    lineHeight: 30,
  }

  it('隣接する段落要素を1グループにまとめる', () => {
    const elements: SlideElement[] = [
      { type: 'paragraph', x: 100, y: 50, width: 800, height: 30,
        runs: [{ text: 'Line 1', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
      { type: 'paragraph', x: 100, y: 90, width: 800, height: 30,
        runs: [{ text: 'Line 2', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    expect(groups).toHaveLength(1)
    expect(groups[0]).toHaveLength(2)
  })

  it('垂直ギャップが大きい要素は別グループに分割する', () => {
    const elements: SlideElement[] = [
      { type: 'paragraph', x: 100, y: 50, width: 800, height: 30,
        runs: [{ text: 'A', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
      { type: 'paragraph', x: 100, y: 200, width: 800, height: 30,
        runs: [{ text: 'B', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    expect(groups).toHaveLength(2)
    expect(groups[0]).toHaveLength(1)
    expect(groups[1]).toHaveLength(1)
  })

  it('X座標が大きく異なる要素は別グループにする', () => {
    const elements: SlideElement[] = [
      { type: 'paragraph', x: 100, y: 50, width: 400, height: 30,
        runs: [{ text: 'Left', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
      { type: 'paragraph', x: 600, y: 80, width: 400, height: 30,
        runs: [{ text: 'Right', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    expect(groups).toHaveLength(2)
  })

  it('幅が異なる要素は別グループにする', () => {
    const elements: SlideElement[] = [
      { type: 'paragraph', x: 100, y: 50, width: 800, height: 30,
        runs: [{ text: 'Wide', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
      { type: 'paragraph', x: 100, y: 80, width: 400, height: 30,
        runs: [{ text: 'Narrow', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    expect(groups).toHaveLength(2)
  })

  it('テーブル・画像などグルーピング対象外の要素はそのまま', () => {
    const elements: SlideElement[] = [
      { type: 'paragraph', x: 100, y: 50, width: 800, height: 30,
        runs: [{ text: 'Para', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
      { type: 'image', x: 100, y: 80, width: 800, height: 200,
        src: 'data:image/png;base64,', naturalWidth: 800, naturalHeight: 200 },
      { type: 'paragraph', x: 100, y: 290, width: 800, height: 30,
        runs: [{ text: 'After image', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    expect(groups).toHaveLength(3)
  })

  it('heading は常に独立したグループとなる', () => {
    const elements: SlideElement[] = [
      { type: 'heading', x: 100, y: 50, width: 800, height: 40, level: 1,
        runs: [{ text: 'Title', color: 'rgb(0,0,0)', fontSize: 36 }], style: baseStyle },
      { type: 'paragraph', x: 100, y: 95, width: 800, height: 30,
        runs: [{ text: 'Body', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
      { type: 'list', x: 100, y: 130, width: 800, height: 60, ordered: false,
        items: [{ text: 'Item 1', level: 0, runs: [{ text: 'Item 1', color: 'rgb(0,0,0)', fontSize: 24 }] }],
        style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    // heading は常に独立; list もグルーピング対象外で独立
    expect(groups).toHaveLength(3)
    expect(groups[0]).toHaveLength(1) // heading only
    expect(groups[0][0].type).toBe('heading')
    expect(groups[1]).toHaveLength(1) // paragraph
    expect(groups[2]).toHaveLength(1) // list
  })

  it('borderBottom 付き heading はグルーピングしない', () => {
    const elements: SlideElement[] = [
      { type: 'heading', x: 100, y: 50, width: 800, height: 40, level: 1,
        runs: [{ text: 'Title', color: 'rgb(0,0,0)', fontSize: 36 }], style: baseStyle,
        borderBottom: { width: 2, color: 'rgb(0,0,0)' } },
      { type: 'paragraph', x: 100, y: 95, width: 800, height: 30,
        runs: [{ text: 'Body', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    expect(groups).toHaveLength(2)
  })

  it('borderLeft 付き blockquote はグルーピングしない', () => {
    const elements: SlideElement[] = [
      { type: 'blockquote', x: 100, y: 50, width: 800, height: 40,
        runs: [{ text: 'Quote', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle,
        borderLeft: { width: 4, color: 'rgb(128,128,128)' } },
      { type: 'paragraph', x: 100, y: 95, width: 800, height: 30,
        runs: [{ text: 'Body', color: 'rgb(0,0,0)', fontSize: 24 }], style: baseStyle },
    ]
    const groups = groupAdjacentTextElements(elements)
    expect(groups).toHaveLength(2)
  })

  it('空配列は空配列を返す', () => {
    expect(groupAdjacentTextElements([])).toEqual([])
  })
})

// ── Table cell paragraph model ────────────────────────────────────────
describe('placeElement – table cell paragraphs', () => {
  const makeSlide = () => ({
    addText: jest.fn(),
    addShape: jest.fn(),
    addImage: jest.fn(),
    addTable: jest.fn(),
    addNotes: jest.fn(),
  })
  const baseStyle = {
    color: 'rgb(0,0,0)',
    backgroundColor: 'rgba(0,0,0,0)',
    fontSize: 16,
    fontFamily: 'Arial',
    fontWeight: 400,
    textAlign: 'left',
    borderColor: 'rgba(0,0,0,0)',
  }

  it('paragraphs が存在する場合 breakLine で段落分離される', () => {
    const cell: TableCell = {
      text: 'Line1\nLine2',
      runs: [
        { text: 'Line1', color: 'rgb(0,0,0)', fontSize: 16 },
        { text: '', breakLine: true },
        { text: 'Line2', color: 'rgb(0,0,0)', fontSize: 16 },
      ],
      paragraphs: [
        { runs: [{ text: 'Line1', color: 'rgb(0,0,0)', fontSize: 16 }] },
        { runs: [{ text: 'Line2', color: 'rgb(0,0,0)', fontSize: 16 }] },
      ],
      isHeader: false,
      style: baseStyle,
    }
    const el: SlideElement = {
      type: 'table',
      x: 0, y: 0, width: 800, height: 200,
      rows: [{ cells: [cell] }],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const rows = (mockSlide.addTable as jest.Mock).mock.calls[0][0] as any[][]
    const textArray = rows[0][0].text
    // paragraph model: each paragraph's last run carries breakLine except final
    expect(textArray).toHaveLength(2)
    expect(textArray[0].text).toBe('Line1')
    expect(textArray[0].options.breakLine).toBe(true)
    expect(textArray[1].text).toBe('Line2')
    expect(textArray[1].options.breakLine).toBeUndefined()
  })

  it('paragraphs が空の場合は runs にフォールバックする', () => {
    const cell: TableCell = {
      text: 'Simple',
      runs: [
        { text: 'Simple', color: 'rgb(0,0,0)', fontSize: 16 },
      ],
      paragraphs: [],
      isHeader: false,
      style: baseStyle,
    }
    const el: SlideElement = {
      type: 'table',
      x: 0, y: 0, width: 800, height: 200,
      rows: [{ cells: [cell] }],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const rows = (mockSlide.addTable as jest.Mock).mock.calls[0][0] as any[][]
    const textArray = rows[0][0].text
    expect(textArray).toHaveLength(1)
    expect(textArray[0].text).toBe('Simple')
  })

  it('3段落のテーブルセルで正しく分離される', () => {
    const cell: TableCell = {
      text: 'A\nB\nC',
      runs: [
        { text: 'A', color: 'rgb(0,0,0)', fontSize: 16 },
        { text: '', breakLine: true },
        { text: 'B', color: 'rgb(0,0,0)', fontSize: 16 },
        { text: '', breakLine: true },
        { text: 'C', color: 'rgb(0,0,0)', fontSize: 16 },
      ],
      paragraphs: [
        { runs: [{ text: 'A', color: 'rgb(0,0,0)', fontSize: 16 }] },
        { runs: [{ text: 'B', color: 'rgb(0,0,0)', fontSize: 16 }] },
        { runs: [{ text: 'C', color: 'rgb(0,0,0)', fontSize: 16 }] },
      ],
      isHeader: false,
      style: baseStyle,
    }
    const el: SlideElement = {
      type: 'table',
      x: 0, y: 0, width: 800, height: 200,
      rows: [{ cells: [cell] }],
      style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
    }

    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const rows = (mockSlide.addTable as jest.Mock).mock.calls[0][0] as any[][]
    const textArray = rows[0][0].text
    // 3 paragraphs → 3 runs: A(breakLine) + B(breakLine) + C
    expect(textArray).toHaveLength(3)
    expect(textArray[0].text).toBe('A')
    expect(textArray[0].options.breakLine).toBe(true)
    expect(textArray[1].text).toBe('B')
    expect(textArray[1].options.breakLine).toBe(true)
    expect(textArray[2].text).toBe('C')
    expect(textArray[2].options.breakLine).toBeUndefined()
  })
})

// ── Code block syntax highlighting (A1) ───────────────────────────────
describe('placeElement – code block syntax highlighting', () => {
  const makeSlide = () => ({
    addText: jest.fn(),
    addShape: jest.fn(),
    addImage: jest.fn(),
    addTable: jest.fn(),
    addNotes: jest.fn(),
  })

  it('syntax-highlighted runs を使って色付きテキストを出力する', () => {
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 800, height: 200,
      text: 'const x = 1',
      language: 'javascript',
      runs: [
        { text: 'const', color: 'rgb(199,146,234)', fontSize: 14 },
        { text: ' x ', color: 'rgb(200,200,200)', fontSize: 14 },
        { text: '=', color: 'rgb(137,221,255)', fontSize: 14 },
        { text: ' 1', color: 'rgb(247,140,108)', fontSize: 14 },
      ],
      style: { color: 'rgb(200,200,200)', fontSize: 14, fontFamily: 'monospace', fontWeight: 400, textAlign: 'left', lineHeight: 20, backgroundColor: 'rgb(40,44,52)' },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    // Code block: single addText with shape+fill (no separate addShape)
    expect(mockSlide.addShape).toHaveBeenCalledTimes(0)
    expect(mockSlide.addText).toHaveBeenCalledTimes(1)

    // Verify shape options include fill for code background
    const textOpts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(textOpts.shape).toBe('rect')
    expect(textOpts.fill).toEqual({ color: '282C34' })

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    expect(Array.isArray(textArg)).toBe(true)
    expect(textArg).toHaveLength(4)
    // First run should be keyword colour
    expect(textArg[0].text).toBe('const')
    expect(textArg[0].options.color).toBe('C792EA')
    // All runs get Courier New
    for (const t of textArg) {
      expect(t.options.fontFace).toBe('Courier New')
    }
  })

  it('runs が空の場合はプレーンテキストにフォールバックする', () => {
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 800, height: 200,
      text: 'plain text',
      language: '',
      runs: [],
      style: { color: 'rgb(200,200,200)', fontSize: 14, fontFamily: 'monospace', fontWeight: 400, textAlign: 'left', lineHeight: 20, backgroundColor: 'rgb(40,44,52)' },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    // Plain text string (not array of runs)
    expect(typeof textArg).toBe('string')
    expect(textArg).toBe('plain text')
  })

  it('breakLine runs で空行を含むコードブロックの行構造が保持される', () => {
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 800, height: 200,
      text: 'a\n\nb',
      language: 'text',
      runs: [
        { text: 'a', color: 'rgb(200,200,200)', fontSize: 14 },
        { text: '', breakLine: true },
        { text: '', breakLine: true },
        { text: 'b', color: 'rgb(200,200,200)', fontSize: 14 },
      ],
      style: { color: 'rgb(200,200,200)', fontSize: 14, fontFamily: 'monospace', fontWeight: 400, textAlign: 'left', lineHeight: 20, backgroundColor: 'rgb(40,44,52)' },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    expect(Array.isArray(textArg)).toBe(true)
    // a + breakLine + breakLine + b = 4 entries
    expect(textArg).toHaveLength(4)
    expect(textArg[0].text).toBe('a')
    expect(textArg[1].options.breakLine).toBe(true)
    expect(textArg[2].options.breakLine).toBe(true)
    expect(textArg[3].text).toBe('b')
  })
})

// ── list-style-type mapping (B1) ──────────────────────────────────────
describe('toListTextProps – listStyleType mapping', () => {
  const baseItem = (text: string, level = 0): import('./types').ListItem => ({
    text,
    level,
    runs: [{ text, color: 'rgb(0,0,0)', fontSize: 24 }],
  })

  it('ordered + lower-alpha → alphaLcPeriod', () => {
    const result = toListTextProps(baseItem('A'), true, false, 'rgb(255,255,255)', false, 1, 'lower-alpha')
    expect((result[0].options?.bullet as any)?.style).toBe('alphaLcPeriod')
  })

  it('ordered + upper-roman → romanUcPeriod', () => {
    const result = toListTextProps(baseItem('I'), true, false, 'rgb(255,255,255)', false, 1, 'upper-roman')
    expect((result[0].options?.bullet as any)?.style).toBe('romanUcPeriod')
  })

  it('ordered + デフォルト(decimal) → arabicPeriod', () => {
    const result = toListTextProps(baseItem('1'), true, false, 'rgb(255,255,255)', false, 1, 'decimal')
    expect((result[0].options?.bullet as any)?.style).toBe('arabicPeriod')
  })

  it('ordered + listStyleType 未指定 → arabicPeriod', () => {
    const result = toListTextProps(baseItem('1'), true, false, 'rgb(255,255,255)', false, 1, undefined)
    expect((result[0].options?.bullet as any)?.style).toBe('arabicPeriod')
  })

  it('unordered + circle → ◦ (25E6)', () => {
    const result = toListTextProps(baseItem('x'), false, false, 'rgb(255,255,255)', false, undefined, 'circle')
    expect((result[0].options?.bullet as any)?.characterCode).toBe('25E6')
  })

  it('unordered + square → ▪ (25AA)', () => {
    const result = toListTextProps(baseItem('x'), false, false, 'rgb(255,255,255)', false, undefined, 'square')
    expect((result[0].options?.bullet as any)?.characterCode).toBe('25AA')
  })

  it('unordered + none → ゼロ幅スペース (200B)', () => {
    const result = toListTextProps(baseItem('x'), false, false, 'rgb(255,255,255)', false, undefined, 'none')
    expect((result[0].options?.bullet as any)?.characterCode).toBe('200B')
  })

  it('unordered + disc (デフォルト) → PptxGenJS デフォルト bullet', () => {
    const result = toListTextProps(baseItem('x'), false, false, 'rgb(255,255,255)', false, undefined, 'disc')
    // Default: bullet is `true` (no characterCode)
    expect(result[0].options?.bullet).toBe(true)
  })

  it('item レベルの listStyleType が親 list の値を上書きする', () => {
    const item = { ...baseItem('i'), listStyleType: 'lower-roman' }
    const result = toListTextProps(item, true, false, 'rgb(255,255,255)', false, 1, 'decimal')
    expect((result[0].options?.bullet as any)?.style).toBe('romanLcPeriod')
  })

  it('ordered list + nested item with listStyleType=circle → unordered bullet (◦)', () => {
    // Simulates <ol><li>...<ul><li>nested</li></ul></li></ol>
    // The nested <li> has listStyleType='circle' from getComputedStyle
    const nestedItem = { ...baseItem('nested', 1), listStyleType: 'circle' }
    const result = toListTextProps(nestedItem, true, false, 'rgb(255,255,255)', false, undefined, undefined)
    // Should be unordered bullet ◦, NOT a numbered item
    expect((result[0].options?.bullet as any)?.characterCode).toBe('25E6')
    expect((result[0].options?.bullet as any)?.type).toBeUndefined()
  })

  it('ordered list + nested item with listStyleType=disc → default unordered bullet', () => {
    const nestedItem = { ...baseItem('nested', 1), listStyleType: 'disc' }
    const result = toListTextProps(nestedItem, true, false, 'rgb(255,255,255)', false, undefined, undefined)
    // disc = PptxGenJS default bullet (true), not numbered
    expect(result[0].options?.bullet).toBe(true)
  })

  it('unordered list + nested item with listStyleType=decimal → ordered numbering', () => {
    // Simulates <ul><li>...<ol><li>nested</li></ol></li></ul>
    const nestedItem = { ...baseItem('nested', 1), listStyleType: 'decimal' }
    const result = toListTextProps(nestedItem, false, false, 'rgb(255,255,255)', false, undefined, undefined)
    expect((result[0].options?.bullet as any)?.type).toBe('number')
    expect((result[0].options?.bullet as any)?.style).toBe('arabicPeriod')
  })
})

// ── Container shape with embedded text (simple card) ──────────────────
describe('placeElement – container shape with embedded text', () => {
  const makeSlide = () => ({
    addText: jest.fn(),
    addShape: jest.fn(),
    addImage: jest.fn(),
    addTable: jest.fn(),
    addNotes: jest.fn(),
  })
  const baseStyle = {
    color: 'rgb(0,0,0)',
    backgroundColor: 'rgba(0,0,0,0)',
    fontSize: 16,
    fontFamily: 'Arial',
    fontWeight: 400,
    textAlign: 'left' as const,
    lineHeight: 24,
  }

  it('シンプルな paragraph 1 つを持つ container は addText with shape を使う', () => {
    const el: SlideElement = {
      type: 'container',
      x: 100, y: 50, width: 600, height: 200,
      style: { backgroundColor: 'rgb(230,230,250)', borderRadius: 0 },
      children: [
        {
          type: 'paragraph',
          x: 116, y: 70, width: 568, height: 30,
          runs: [{ text: 'カードのテキスト', color: 'rgb(0,0,0)', fontSize: 16 }],
          style: baseStyle,
        },
      ],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    // 埋め込みパスでは addShape を呼ばず addText with shape を使う
    expect(mockSlide.addShape).not.toHaveBeenCalled()
    expect(mockSlide.addText).toHaveBeenCalledTimes(1)

    const [, opts] = (mockSlide.addText as jest.Mock).mock.calls[0]
    expect(opts.shape).toBeDefined()
    expect(opts.fill).toEqual({ color: 'E6E6FA' })
    expect(opts.autoFit).toBe(false)
    expect(opts.wrap).toBe(true)
  })

  it('複数 paragraph を持つ container は runs を統合して addText 1 回を呼ぶ', () => {
    const el: SlideElement = {
      type: 'container',
      x: 100, y: 50, width: 600, height: 300,
      style: { backgroundColor: 'rgb(255,255,255)', borderWidth: 1, borderColor: 'rgb(200,200,200)' },
      children: [
        {
          type: 'paragraph',
          x: 116, y: 66, width: 568, height: 30,
          runs: [{ text: '行 1', color: 'rgb(0,0,0)', fontSize: 16 }],
          style: baseStyle,
        },
        {
          type: 'paragraph',
          x: 116, y: 106, width: 568, height: 30,
          runs: [{ text: '行 2', color: 'rgb(0,0,0)', fontSize: 16 }],
          style: baseStyle,
        },
      ],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    expect(mockSlide.addShape).not.toHaveBeenCalled()
    expect(mockSlide.addText).toHaveBeenCalledTimes(1)
    const runsArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    // 2 paragraphs → runs of first + breakLine + runs of second
    const texts = (runsArg as any[]).filter((r: any) => !r.options?.breakLine)
    const breaks = (runsArg as any[]).filter((r: any) => r.options?.breakLine)
    expect(texts).toHaveLength(2)
    expect(breaks).toHaveLength(1)
  })

  it('borderLeft 付き container は embedded path で統合され、バーは別 rect として描画される', () => {
    const el: SlideElement = {
      type: 'container',
      x: 100, y: 50, width: 600, height: 200,
      style: {
        backgroundColor: 'rgb(245,245,245)',
        borderLeft: { width: 4, color: 'rgb(0,120,215)' },
      },
      children: [
        {
          type: 'paragraph',
          x: 120, y: 70, width: 560, height: 30,
          runs: [{ text: 'note', color: 'rgb(0,0,0)', fontSize: 16 }],
          style: baseStyle,
        },
      ],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    // Embedded path: addText with shape (text+bg 統合) + addShape (border-left bar)
    expect(mockSlide.addText).toHaveBeenCalledTimes(1)
    const [, textOpts] = (mockSlide.addText as jest.Mock).mock.calls[0]
    expect(textOpts.shape).toBe('rect')
    expect(textOpts.fill).toEqual({ color: 'F5F5F5' })

    // Border-left bar drawn as a thin rect
    expect(mockSlide.addShape).toHaveBeenCalledTimes(1)
    const [, barOpts] = (mockSlide.addShape as jest.Mock).mock.calls[0]
    expect(barOpts.fill).toEqual({ color: '0078D7' })
  })

  it('badge/chip runs 付き container は単一オブジェクト (shape+text) として出力する', () => {
    const el: SlideElement = {
      type: 'container',
      x: 200, y: 100, width: 120, height: 32,
      style: { backgroundColor: 'rgb(0,120,215)' },
      runs: [{ text: 'NEW', color: 'rgb(255,255,255)', fontSize: 12 }],
      children: [],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    // badge は addText with shape (単一オブジェクト)
    expect(mockSlide.addShape).toHaveBeenCalledTimes(0)
    expect(mockSlide.addText).toHaveBeenCalledTimes(1)
    const [, opts] = (mockSlide.addText as jest.Mock).mock.calls[0]
    expect(opts.shape).toBe('rect')
    expect(opts.fill).toEqual({ color: '0078D7' })
  })

  it('非表示コンテナ (background なし) は埋め込みパスに入らず子要素だけ描画される', () => {
    const el: SlideElement = {
      type: 'container',
      x: 100, y: 50, width: 600, height: 200,
      style: { backgroundColor: 'rgba(0,0,0,0)' },
      children: [
        {
          type: 'paragraph',
          x: 116, y: 70, width: 568, height: 30,
          runs: [{ text: 'テキスト', color: 'rgb(0,0,0)', fontSize: 16 }],
          style: baseStyle,
        },
      ],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    expect(mockSlide.addShape).not.toHaveBeenCalled()
    // addText は子 paragraph の 1 回。shape プロパティなし
    expect(mockSlide.addText).toHaveBeenCalledTimes(1)
    const [, opts] = (mockSlide.addText as jest.Mock).mock.calls[0]
    expect(opts.shape).toBeUndefined()
  })

  it('roundRect コンテナは rectRadius 付きで shape オプションが設定される', () => {
    const el: SlideElement = {
      type: 'container',
      x: 100, y: 50, width: 600, height: 200,
      style: { backgroundColor: 'rgb(200,220,255)', borderRadius: 12 },
      children: [
        {
          type: 'paragraph',
          x: 116, y: 70, width: 568, height: 30,
          runs: [{ text: '丸角カード', color: 'rgb(0,0,0)', fontSize: 16 }],
          style: baseStyle,
        },
      ],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    expect(mockSlide.addShape).not.toHaveBeenCalled()
    const [, opts] = (mockSlide.addText as jest.Mock).mock.calls[0]
    expect(opts.shape).toBe('roundRect')
    expect(opts.rectRadius).toBeGreaterThan(0)
  })
})

// ── associateContainerText – 空間包含マッチング ────────────────────────
describe('associateContainerText', () => {
  const SLIDE_W = 1280
  const SLIDE_H = 720

  const makeContainer = (x: number, y: number, w: number, h: number, bg = 'rgb(246,248,250)'): SlideElement => ({
    type: 'container',
    x, y, width: w, height: h,
    style: { backgroundColor: bg, borderRadius: 0 },
    children: [],
  })

  const makeParagraph = (x: number, y: number, w: number, h: number): SlideElement => ({
    type: 'paragraph',
    x, y, width: w, height: h,
    runs: [{ text: 'テキスト', color: 'rgb(0,0,0)', fontSize: 16 }],
    style: { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left', lineHeight: 24 },
  })

  it('テキストがコンテナ内に収まっている場合は関連付けられる', () => {
    const container = makeContainer(100, 50, 600, 200)
    const text = makeParagraph(116, 70, 568, 30)
    const result = associateContainerText([container, text], SLIDE_W, SLIDE_H)

    expect(result.size).toBe(1)
    const texts = result.get(container as any)
    expect(texts).toBeDefined()
    expect(texts).toContain(text)
  })

  it('テキストがコンテナの外にある場合は関連付けられない', () => {
    const container = makeContainer(100, 50, 600, 200)
    const textOutside = makeParagraph(50, 300, 200, 30)  // y: 300 > 50+200=250
    const result = associateContainerText([container, textOutside], SLIDE_W, SLIDE_H)

    expect(result.size).toBe(0)
  })

  it('透明な背景のコンテナ（レイアウトラッパー）は除外される', () => {
    const transparent = makeContainer(0, 0, SLIDE_W, SLIDE_H, 'rgba(0,0,0,0)')
    const text = makeParagraph(100, 100, 400, 30)
    const result = associateContainerText([transparent, text], SLIDE_W, SLIDE_H)

    expect(result.size).toBe(0)
  })

  it('スライド全体を覆うコンテナ（背景）は除外される', () => {
    // 幅・高さがスライドの90%以上を超える場合は除外
    const fullSlide = makeContainer(0, 0, SLIDE_W, SLIDE_H, 'rgb(200,200,200)')
    const text = makeParagraph(100, 100, 400, 30)
    const result = associateContainerText([fullSlide, text], SLIDE_W, SLIDE_H)

    expect(result.size).toBe(0)
  })

  it('複数コンテナがある場合、テキストは最小（最内側）コンテナに関連付けられる', () => {
    const outer = makeContainer(50, 50, 700, 400, 'rgb(230,230,230)')
    const inner = makeContainer(100, 80, 300, 150, 'rgb(246,248,250)')
    const text = makeParagraph(110, 90, 280, 30)  // inner の中にある
    const result = associateContainerText([outer, inner, text], SLIDE_W, SLIDE_H)

    // text は inner に関連付けられる（面積が小さい方）
    expect(result.get(inner as any)).toContain(text)
    // outer には関連付けられない（inner が先に取る）
    expect(result.get(outer as any) ?? []).not.toContain(text)
  })

  it('コンテナ内に複数テキスト要素がある場合はすべて関連付けられ、y 座標でソートされる', () => {
    const container = makeContainer(100, 50, 600, 300)
    const text1 = makeParagraph(116, 200, 568, 30)  // y: 200
    const text2 = makeParagraph(116, 80, 568, 30)   // y: 80  (上にある)
    const result = associateContainerText([container, text1, text2], SLIDE_W, SLIDE_H)

    const texts = result.get(container as any)!
    expect(texts).toHaveLength(2)
    // y 座標昇順でソートされている
    expect(texts[0]).toBe(text2)
    expect(texts[1]).toBe(text1)
  })

  it('コンテナ要素のみで、テキストが一切内包されていない場合は結果 Map に含まれない', () => {
    const container = makeContainer(100, 50, 600, 200)
    const result = associateContainerText([container], SLIDE_W, SLIDE_H)

    expect(result.size).toBe(0)
  })
})

// ── 子レベルのコンテナ・テキスト統合 (badge circle / chat bubble) ──────
describe('placeElement – 子レベルでのコンテナ・テキスト統合', () => {
  const makeSlide = () => ({
    addText: jest.fn(),
    addShape: jest.fn(),
    addImage: jest.fn(),
    addTable: jest.fn(),
    addNotes: jest.fn(),
  })
  const baseStyle = { color: 'rgb(0,0,0)', fontSize: 16, fontFamily: 'Arial', fontWeight: 400, textAlign: 'left' as const, lineHeight: 24 }

  it('透明コンテナ内のバッジ circle + paragraph が 1 オブジェクト (shape+text) に統合される', () => {
    // Simulates: row container (transparent) → [badge circle, badge number, label]
    const el: SlideElement = {
      type: 'container',
      x: 79, y: 326, width: 1123, height: 44,
      style: { backgroundColor: 'rgba(0, 0, 0, 0)' },
      children: [
        {
          type: 'container',
          x: 79, y: 332, width: 32, height: 32,
          style: { backgroundColor: 'rgb(0, 102, 204)', borderRadius: 50 },
          children: [],
        },
        {
          type: 'paragraph',
          x: 79, y: 332, width: 40, height: 32,
          runs: [{ text: '1', color: 'rgb(255,255,255)', fontSize: 14 }],
          style: { ...baseStyle, color: 'rgb(255,255,255)', fontSize: 14 },
        },
        {
          type: 'paragraph',
          x: 121, y: 326, width: 368, height: 44,
          runs: [{ text: 'Label text', color: 'rgb(31,35,40)', fontSize: 14 }],
          style: { ...baseStyle, color: 'rgb(31,35,40)', fontSize: 14 },
        },
      ],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    // 透明コンテナ自体は addShape されない
    // badge circle (shape+text) で 1 回 + label paragraph で 1 回 = addText 2 回
    expect(mockSlide.addShape).toHaveBeenCalledTimes(0)
    expect(mockSlide.addText).toHaveBeenCalledTimes(2)

    // 1 回目: badge circle + number text が統合
    const [, badgeOpts] = (mockSlide.addText as jest.Mock).mock.calls[0]
    expect(badgeOpts.shape).toBeDefined() // roundRect or rect
    expect(badgeOpts.fill).toEqual({ color: '0066CC' })
  })

  it('透明コンテナ内のカード背景 + テキストが 1 オブジェクトに統合される', () => {
    // Simulates: cards wrapper (transparent) → [card bg, card text]
    const el: SlideElement = {
      type: 'container',
      x: 79, y: 348, width: 1123, height: 146,
      style: { backgroundColor: 'rgba(0, 0, 0, 0)' },
      children: [
        {
          type: 'container',
          x: 79, y: 348, width: 556, height: 146,
          style: { backgroundColor: 'rgb(240, 248, 255)' },
          children: [],
        },
        {
          type: 'paragraph',
          x: 79, y: 348, width: 564, height: 146,
          runs: [{ text: 'Card content', color: 'rgb(0,0,0)', fontSize: 16 }],
          style: baseStyle,
        },
      ],
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    // 透明コンテナ: addShape 0 回
    // カード背景 + テキスト統合 = addText 1 回 (with shape)
    expect(mockSlide.addShape).toHaveBeenCalledTimes(0)
    expect(mockSlide.addText).toHaveBeenCalledTimes(1)
    const [, opts] = (mockSlide.addText as jest.Mock).mock.calls[0]
    expect(opts.shape).toBe('rect')
    expect(opts.fill).toEqual({ color: 'F0F8FF' })
  })
})

// ── ADR-45: Code block overflow, list marker size, code indent preservation ──
describe('ADR-45: code block autoFit and overflow containment', () => {
  const makeSlide = () => ({
    addText: jest.fn(),
    addShape: jest.fn(),
    addImage: jest.fn(),
    addTable: jest.fn(),
    addNotes: jest.fn(),
  })

  it('code block shape options include autoFit:false to prevent text overflow', () => {
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 600, height: 200,
      text: '  const x = 1;\n  const y = 2;',
      language: 'typescript',
      runs: [
        { text: '  const', color: 'rgb(199,146,234)', fontSize: 13 },
        { text: ' x = ', color: 'rgb(200,200,200)', fontSize: 13 },
        { text: '1', color: 'rgb(247,140,108)', fontSize: 13 },
        { text: '', breakLine: true },
        { text: '  const', color: 'rgb(199,146,234)', fontSize: 13 },
        { text: ' y = ', color: 'rgb(200,200,200)', fontSize: 13 },
        { text: '2', color: 'rgb(247,140,108)', fontSize: 13 },
      ],
      style: {
        color: 'rgb(200,200,200)', fontSize: 13, fontFamily: 'monospace',
        fontWeight: 400, textAlign: 'left', lineHeight: 18,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    // Code blocks must prevent text from overflowing the shape boundary
    expect(opts.autoFit).toBe(false)
  })

  it('code block shape options include wrap:false to prevent word-wrapping', () => {
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 600, height: 200,
      text: 'const veryLongVariableName = someReallyLongFunctionCall(argumentOne, argumentTwo)',
      language: 'typescript',
      runs: [
        { text: 'const veryLongVariableName = someReallyLongFunctionCall(argumentOne, argumentTwo)', color: 'rgb(200,200,200)', fontSize: 13 },
      ],
      style: {
        color: 'rgb(200,200,200)', fontSize: 13, fontFamily: 'monospace',
        fontWeight: 400, textAlign: 'left', lineHeight: 18,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    // Code blocks must NOT word-wrap — preserves line structure
    expect(opts.wrap).toBe(false)
  })

  it('code block preserves leading whitespace in run text', () => {
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 600, height: 200,
      text: '    indent4\n      indent6',
      language: 'text',
      runs: [
        { text: '    indent4', color: 'rgb(200,200,200)', fontSize: 14 },
        { text: '', breakLine: true },
        { text: '      indent6', color: 'rgb(200,200,200)', fontSize: 14 },
      ],
      style: {
        color: 'rgb(200,200,200)', fontSize: 14, fontFamily: 'monospace',
        fontWeight: 400, textAlign: 'left', lineHeight: 20,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    // Leading spaces must NOT be stripped
    expect(textArg[0].text).toBe('    indent4')
    expect(textArg[2].text).toBe('      indent6')
  })
})

describe('ADR-45: circle bullet marker size proportionality', () => {
  it('circle bullet uses a small glyph character that is visually proportionate to text', () => {
    const item: import('./types').ListItem = {
      text: 'Cat-A val-1',
      level: 1,
      runs: [{ text: 'Cat-A val-1', color: 'rgb(0,0,0)', fontSize: 24 }],
      listStyleType: 'circle',
    }
    const result = toListTextProps(item, false, false, 'rgb(255,255,255)', false, undefined, 'circle')
    const bullet = result[0].options?.bullet as Record<string, any>
    // The character must be a small open-circle glyph, NOT the full-size ○ (U+25CB)
    expect(bullet.characterCode).toBeDefined()
    // Must NOT be U+25CB (WHITE CIRCLE) which is oversized in PowerPoint
    expect(bullet.characterCode).not.toBe('25CB')
    // Acceptable small markers: U+25E6 (white bullet), U+00B7 (middle dot),
    // U+2218 (ring operator), U+2219 (bullet operator)
    const acceptableSmallMarkers = ['25E6', '00B7', '2218', '2219']
    expect(acceptableSmallMarkers).toContain(bullet.characterCode)
  })
})

// ── ADR-46: Detection tests for font size fidelity, shape fill, and text completeness ──
describe('ADR-46: code block fontSize fidelity in PPTX output', () => {
  const makeSlide = () => ({
    addText: jest.fn(),
    addShape: jest.fn(),
    addImage: jest.fn(),
    addTable: jest.fn(),
    addNotes: jest.fn(),
  })

  it('code block text runs use pxToPoints(run.fontSize) — not a default or fallback', () => {
    // If code runs have fontSize 18 (typical Marp code block at 1280x720),
    // the PPTX text run must use 13.5pt (= 18 * 0.75), NOT 16*0.75=12 or 24*0.75=18
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 900, height: 300,
      text: 'const service = new ReservationService();\nservice.book(seat);',
      language: 'typescript',
      runs: [
        { text: 'const', color: 'rgb(199,146,234)', fontSize: 18 },
        { text: ' service = ', color: 'rgb(200,200,200)', fontSize: 18 },
        { text: 'new', color: 'rgb(199,146,234)', fontSize: 18 },
        { text: ' ReservationService();', color: 'rgb(200,200,200)', fontSize: 18 },
        { text: '', breakLine: true },
        { text: 'service.book(seat);', color: 'rgb(200,200,200)', fontSize: 18 },
      ],
      style: {
        color: 'rgb(200,200,200)', fontSize: 18, fontFamily: 'monospace',
        fontWeight: 400, textAlign: 'left', lineHeight: 26,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    // Every non-breakLine run must have fontSize = 18 * 0.75 = 13.5
    const textRuns = textArg.filter((r: any) => r.text !== '' || !r.options?.breakLine)
    for (const run of textRuns) {
      if (run.options?.breakLine) continue
      expect(run.options.fontSize).toBeCloseTo(13.5, 1)
    }
  })

  it('code block text runs respect per-run fontSize when runs have varying sizes', () => {
    // Syntax-highlighted code may have different fontSizes per span (e.g. superscript annotations)
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 900, height: 200,
      text: 'main()  // entry',
      language: 'typescript',
      runs: [
        { text: 'main()', color: 'rgb(130,170,255)', fontSize: 14 },
        { text: '  // entry', color: 'rgb(100,100,100)', fontSize: 11 },
      ],
      style: {
        color: 'rgb(200,200,200)', fontSize: 14, fontFamily: 'monospace',
        fontWeight: 400, textAlign: 'left', lineHeight: 20,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    // First run: 14px → 10.5pt
    expect(textArg[0].options.fontSize).toBeCloseTo(10.5, 1)
    // Second run: 11px → 8.25pt
    expect(textArg[1].options.fontSize).toBeCloseTo(8.25, 1)
  })

  it('code block shape uses fill color from element style.backgroundColor', () => {
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 900, height: 200,
      text: '.button { color: blue; }',
      language: 'css',
      runs: [
        { text: '.button { color: blue; }', color: 'rgb(200,200,200)', fontSize: 14 },
      ],
      style: {
        color: 'rgb(200,200,200)', fontSize: 14, fontFamily: 'monospace',
        fontWeight: 400, textAlign: 'left', lineHeight: 20,
        backgroundColor: 'rgb(246,248,250)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    // The shape must have a fill matching the code block background
    expect(opts.fill).toBeDefined()
    expect(opts.fill.color).toBe('F6F8FA')
  })

  it('code block fontSize falls back to element style when run has no fontSize', () => {
    // When runs are extracted without per-run fontSize (undefined),
    // the builder must use el.style.fontSize as fallback
    const el: SlideElement = {
      type: 'code',
      x: 50, y: 100, width: 900, height: 200,
      text: 'hello world',
      language: 'text',
      runs: [
        { text: 'hello world', color: 'rgb(200,200,200)' } as any, // no fontSize
      ],
      style: {
        color: 'rgb(200,200,200)', fontSize: 15, fontFamily: 'monospace',
        fontWeight: 400, textAlign: 'left', lineHeight: 22,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    // Fallback: 15px → 11.25pt
    expect(textArg[0].options.fontSize).toBeCloseTo(11.25, 1)
  })
})

// ── ADR-46 Gate 2: Output validation — text must physically fit in shape ──
// This is a COMPLETELY DIFFERENT detection angle from the extraction-level tests.
// It validates that the PPTX output is physically sensible regardless of how the
// data was extracted. Catches issues even when dom-walker auto-scaling detection fails.
describe('ADR-46 Gate 2: code block fontSize must not cause vertical overflow in PPTX shape', () => {
  const makeSlide = () => ({
    addText: jest.fn(),
    addShape: jest.fn(),
    addImage: jest.fn(),
    addTable: jest.fn(),
    addNotes: jest.fn(),
  })

  it('20-line code block at 24.65px fontSize in 500px-tall box: fontSize must be scaled to fit', () => {
    // Simulates Slide 90: BookingService class with ~20 lines
    // Without fix: fontSize=24.65px → 18.49pt, 20 lines * 18.49pt * 1.2 = 443pt
    // Box height = 500px → 375pt (= 5.2 inches) — text overflows!
    // With fix: fontSize must be reduced so all 20 lines fit
    const numLines = 20
    const runs: any[] = []
    for (let i = 0; i < numLines; i++) {
      if (i > 0) runs.push({ text: '', breakLine: true })
      runs.push({ text: `  line ${i + 1}: code content here`, color: 'rgb(200,200,200)', fontSize: 24.65 })
    }

    const el: SlideElement = {
      type: 'code',
      x: 50, y: 120, width: 1100, height: 500,
      text: runs.filter(r => !r.breakLine).map(r => r.text).join('\n'),
      language: 'typescript',
      runs,
      style: {
        color: 'rgb(200,200,200)', fontSize: 24.65, fontFamily: 'Courier New',
        fontWeight: 400, textAlign: 'left', lineHeight: 28.35,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    const textRuns = textArg.filter((r: any) => !r.options?.breakLine)

    // The shape height in inches
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    const shapeHeightInches = opts.h ?? (500 / 96)
    const shapeHeightPt = shapeHeightInches * 72

    // Verify: total text height must fit within shape
    // fontSize (pt) * numLines * lineSpacing <= shapeHeight (pt)
    // Allow 1pt tolerance for floating-point arithmetic
    const actualFontSizePt = textRuns[0].options.fontSize
    const estimatedTotalHeight = actualFontSizePt * numLines * 1.2
    expect(estimatedTotalHeight).toBeLessThanOrEqual(shapeHeightPt + 1)

    // Also verify fontSize is smaller than the raw unscaled value
    const rawFontSizePt = 24.65 * 0.75 // = 18.49pt
    expect(actualFontSizePt).toBeLessThan(rawFontSizePt)
  })

  it('3-line code block at 24.65px fontSize in 200px-tall box: fits without scaling', () => {
    // 3 lines * 18.49pt * 1.2 = 66.6pt needed
    // Box height 200px → 150pt available — fits fine!
    const runs: any[] = [
      { text: 'const x = 1;', color: 'rgb(200,200,200)', fontSize: 24.65 },
      { text: '', breakLine: true },
      { text: 'const y = 2;', color: 'rgb(200,200,200)', fontSize: 24.65 },
      { text: '', breakLine: true },
      { text: 'const z = 3;', color: 'rgb(200,200,200)', fontSize: 24.65 },
    ]

    const el: SlideElement = {
      type: 'code',
      x: 50, y: 120, width: 1100, height: 200,
      text: 'const x = 1;\nconst y = 2;\nconst z = 3;',
      language: 'typescript',
      runs,
      style: {
        color: 'rgb(200,200,200)', fontSize: 24.65, fontFamily: 'Courier New',
        fontWeight: 400, textAlign: 'left', lineHeight: 28.35,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    const textRuns = textArg.filter((r: any) => !r.options?.breakLine)

    // No scaling needed — font stays at 18.49pt
    const actualFontSizePt = textRuns[0].options.fontSize
    expect(actualFontSizePt).toBeCloseTo(24.65 * 0.75, 1) // 18.49pt unchanged
  })

  it('25-line code block overflowing slide (height > 720 - y): fontSize capped to slide area', () => {
    // Simulates a code block that extends beyond the slide boundary
    // 25 lines at fontSize 24.65px, y=120 → available slide height = 720-120 = 600px
    // Slide builder clamps h to slide area, but font must also be scaled
    const numLines = 25
    const runs: any[] = []
    for (let i = 0; i < numLines; i++) {
      if (i > 0) runs.push({ text: '', breakLine: true })
      runs.push({ text: `    line ${i + 1}`, color: 'rgb(200,200,200)', fontSize: 24.65 })
    }

    const el: SlideElement = {
      type: 'code',
      x: 50, y: 120, width: 1100, height: 800, // extends beyond slide!
      text: runs.filter(r => !r.breakLine).map(r => r.text).join('\n'),
      language: 'typescript',
      runs,
      style: {
        color: 'rgb(200,200,200)', fontSize: 24.65, fontFamily: 'Courier New',
        fontWeight: 400, textAlign: 'left', lineHeight: 28.35,
        backgroundColor: 'rgb(40,44,52)',
      },
    }
    const mockSlide = makeSlide()
    placeElement(mockSlide as any, el, 1280, 720, 'rgb(255,255,255)', false)

    const textArg = (mockSlide.addText as jest.Mock).mock.calls[0][0]
    const textRuns = textArg.filter((r: any) => !r.options?.breakLine)

    // Shape is clamped to slide area: max h = (720 - 120) / 96 inches
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    const shapeHeightInches = opts.h
    const shapeHeightPt = shapeHeightInches * 72

    // Verify: text fits within clamped shape (1pt tolerance for floating-point)
    const actualFontSizePt = textRuns[0].options.fontSize
    const estimatedTotalHeight = actualFontSizePt * numLines * 1.2
    expect(estimatedTotalHeight).toBeLessThanOrEqual(shapeHeightPt + 1)
    // Must be smaller than raw value
    expect(actualFontSizePt).toBeLessThan(24.65 * 0.75)
  })
})
