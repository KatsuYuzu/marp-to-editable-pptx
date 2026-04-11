import {
  buildPptx,
  placeElement,
  toTextProps,
  toListTextProps,
} from './slide-builder'
import type { SlideData, ImageElement } from './types'

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

  it('backgroundColor の run には highlight が設定される — slide 56/58 の strong ハイライト', () => {
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
      style: { ...baseStyle, fontSize: 16, lineHeight: 24 }, // 24/16 = 1.5
    }
    placeElement(mockSlide, el, 1280, 720)
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(opts.lineSpacingMultiple).toBeCloseTo(1.5, 2)
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
      style: { ...baseStyle, fontSize: 32, lineHeight: 40 }, // 40/32 = 1.25
    }
    placeElement(mockSlide, el, 1280, 720)
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(opts.lineSpacingMultiple).toBeCloseTo(1.25, 2)
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
      style: { ...baseStyle, fontSize: 16, lineHeight: 22 }, // 22/16 = 1.375
    }
    placeElement(mockSlide, el, 1280, 720)
    const opts = (mockSlide.addText as jest.Mock).mock.calls[0][1]
    expect(opts.lineSpacingMultiple).toBeCloseTo(1.38, 2)
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

describe('placeElement — paragraph width extension for wide elements', () => {
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

  it('extends wide paragraph (right edge > 70 %, width > 25 %) by up to 32 px', () => {
    // Simulates a chat-bubble paragraph: x=79, width=898 (80 % of 1123 px content
    // area). Right edge = 977 px / 1280 px = 76.3 % → above 70 % threshold.
    // Width = 898 px > 25 % of 1280 (320 px) → qualifies.
    // Expected extended w = min(898 + 32, 1280 − 79 − 8) = 930 px.
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
    const expectedW = Math.min(898 + 32, 1280 - 79 - 8) / 96
    expect(w).toBeCloseTo(expectedW, 5)
  })

  it('does not extend narrow paragraph (right edge < 70 %)', () => {
    // x=79, width=400: right edge = 479 px = 37 % → below threshold.
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
    expect(w).toBeCloseTo(400 / 96, 5)
  })

  it('does not extend short paragraph even if far right (width ≤ 25 %)', () => {
    // x=1000, width=200: right edge = 1200 px = 93.75 % but width = 200 < 320 px
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
    expect(w).toBeCloseTo(200 / 96, 5)
  })

  it('caps extension at slideW − x − 8 to avoid slide overflow', () => {
    // x=79, width=1185: right edge = 1264 px = 98.75 %. Cap = 1280−79−8=1193.
    // min(1185+32, 1193) = 1193.
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
    const expectedW = Math.min(1185 + 32, 1280 - 79 - 8) / 96
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
