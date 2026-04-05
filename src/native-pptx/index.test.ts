import { generateNativePptx } from './index'

// Mock node:fs/promises so tests never touch the real file system
const mockAccess = jest.fn().mockResolvedValue(undefined)
const mockReadFile = jest.fn().mockResolvedValue(Buffer.alloc(0))
const mockWriteFile = jest.fn().mockResolvedValue(undefined)
jest.mock('node:fs/promises', () => ({
  access: (...args: any[]) => mockAccess(...args),
  readFile: (...args: any[]) => mockReadFile(...args),
  writeFile: (...args: any[]) => mockWriteFile(...args),
}))

// Mock puppeteer-core
const mockClose = jest.fn()
const mockSetViewport = jest.fn()
const mockGoto = jest.fn()
const mockEvaluate = jest.fn()
const mockAddScriptTag = jest.fn()
const mockAddStyleTag = jest.fn()
const mockScreenshot = jest.fn().mockResolvedValue(Buffer.from('fakepng'))
const mockNewPage = jest.fn().mockResolvedValue({
  setViewport: mockSetViewport,
  goto: mockGoto,
  evaluate: mockEvaluate,
  addScriptTag: mockAddScriptTag,
  addStyleTag: mockAddStyleTag,
  screenshot: mockScreenshot,
})
const mockLaunch = jest.fn().mockResolvedValue({
  newPage: mockNewPage,
  close: mockClose,
})

jest.mock('puppeteer-core', () => ({
  __esModule: true,
  default: { launch: (...args: any[]) => mockLaunch(...args) },
}))

// Mock slide-builder
const mockWrite = jest.fn().mockResolvedValue(new ArrayBuffer(8))
const mockBuildPptx = jest.fn().mockReturnValue({ write: mockWrite })
jest.mock('./slide-builder', () => ({
  buildPptx: (...args: any[]) => mockBuildPptx(...args),
}))

// Mock the generated dom-walker script (string constant)
jest.mock('./dom-walker-script.generated', () => ({
  DOM_WALKER_SCRIPT: 'globalThis.extractSlides = function() { return []; };',
}))

describe('generateNativePptx', () => {
  beforeEach(() => {
    jest.clearAllMocks()
    // デフォルト: ファイルは存在する、スクリーンショットは成功
    mockAccess.mockResolvedValue(undefined)
    mockScreenshot.mockResolvedValue(Buffer.from('fakepng'))
    mockEvaluate.mockResolvedValue([
      {
        width: 1280,
        height: 720,
        background: 'rgb(255,255,255)',
        backgroundImages: [],
        elements: [],
        notes: '',
      },
    ])
  })

  it('launches browser, loads HTML, extracts DOM, and returns PPTX buffer', async () => {
    const result = await generateNativePptx({
      htmlPath: '/tmp/test.html',
      browserPath: '/usr/bin/chrome',
    })

    // puppeteer launched with correct browser path
    expect(mockLaunch).toHaveBeenCalledWith(
      expect.objectContaining({
        executablePath: '/usr/bin/chrome',
        headless: true,
      }),
    )

    // HTML was loaded via file:// URL to resolve relative paths correctly
    expect(mockGoto).toHaveBeenCalledWith(
      expect.stringMatching(/^file:\/\/\/.*test\.html$/),
      expect.objectContaining({ waitUntil: 'networkidle0' }),
    )

    // Bespoke UI elements (OSC overlay, note panels) were hidden via CSS
    // so they don't appear in Puppeteer background screenshots
    expect(mockAddStyleTag).toHaveBeenCalledWith(
      expect.objectContaining({
        content: expect.stringContaining('bespoke-marp-osc'),
      }),
    )

    // DOM walker script was injected via addScriptTag
    expect(mockAddScriptTag).toHaveBeenCalledWith(
      expect.objectContaining({ content: expect.any(String) }),
    )

    // extractSlides was called via page.evaluate
    expect(mockEvaluate).toHaveBeenCalled()

    // buildPptx was called with extracted slides
    expect(mockBuildPptx).toHaveBeenCalledWith([
      expect.objectContaining({ width: 1280, height: 720 }),
    ])

    // Returns a Buffer
    expect(result).toBeInstanceOf(Buffer)
  })

  it('uses specified viewport size', async () => {
    await generateNativePptx({
      htmlPath: '/tmp/test.html',
      browserPath: '/usr/bin/chrome',
      width: 1920,
      height: 1080,
    })

    expect(mockSetViewport).toHaveBeenCalledWith({
      width: 1920,
      height: 1080,
    })
  })

  it('defaults to 1280x720 viewport', async () => {
    await generateNativePptx({
      htmlPath: '/tmp/test.html',
      browserPath: '/usr/bin/chrome',
    })

    expect(mockSetViewport).toHaveBeenCalledWith({
      width: 1280,
      height: 720,
    })
  })

  it('closes browser after completion', async () => {
    await generateNativePptx({
      htmlPath: '/tmp/test.html',
      browserPath: '/usr/bin/chrome',
    })

    expect(mockClose).toHaveBeenCalled()
  })

  it('closes browser even on error', async () => {
    mockEvaluate.mockRejectedValue(new Error('DOM extraction failed'))

    await expect(
      generateNativePptx({
        htmlPath: '/tmp/test.html',
        browserPath: '/usr/bin/chrome',
      }),
    ).rejects.toThrow('DOM extraction failed')

    expect(mockClose).toHaveBeenCalled()
  })

  describe('欠損画像ファイルの処理', () => {
    it('コンテンツ画像が欠損している場合、その領域をスクリーンショットして src を置き換える', async () => {
      // file:///missing.png は存在しないとみなす
      mockAccess.mockImplementation(async (p: string) => {
        if (p.includes('missing')) {
          throw Object.assign(new Error('ENOENT'), { code: 'ENOENT' })
        }
      })
      mockEvaluate.mockResolvedValue([
        {
          width: 1280,
          height: 720,
          background: 'rgb(255,255,255)',
          backgroundImages: [],
          elements: [
            {
              type: 'image',
              src: 'file:///missing.png',
              naturalWidth: 200,
              naturalHeight: 150,
              x: 100,
              y: 50,
              width: 200,
              height: 150,
            },
          ],
          notes: '',
        },
      ])

      await generateNativePptx({
        htmlPath: '/tmp/test.html',
        browserPath: '/usr/bin/chrome',
      })

      // 欠損画像の領域がスクリーンショットされた
      expect(mockScreenshot).toHaveBeenCalledWith(
        expect.objectContaining({ type: 'png' }),
      )

      // buildPptx にはスクリーンショットの data URL が渡された（元の file:// URL ではない）
      const slides = mockBuildPptx.mock.calls[0][0]
      expect(slides[0].elements[0].src).toMatch(/^data:image\/png;base64,/)
    })

    it('背景画像が欠損している場合、backgroundImages から除去してスライド背景色にフォールバックさせる', async () => {
      mockAccess.mockImplementation(async (p: string) => {
        if (p.includes('missing-bg')) {
          throw Object.assign(new Error('ENOENT'), { code: 'ENOENT' })
        }
      })
      mockEvaluate.mockResolvedValue([
        {
          width: 1280,
          height: 720,
          background: 'rgb(0,0,0)',
          backgroundImages: [
            {
              url: 'file:///missing-bg.png',
              x: 0,
              y: 0,
              width: 1280,
              height: 720,
            },
          ],
          elements: [],
          notes: '',
        },
      ])

      await generateNativePptx({
        htmlPath: '/tmp/test.html',
        browserPath: '/usr/bin/chrome',
      })

      // buildPptx には backgroundImages が空で渡された
      const slides = mockBuildPptx.mock.calls[0][0]
      expect(slides[0].backgroundImages).toHaveLength(0)
    })
  })
})
