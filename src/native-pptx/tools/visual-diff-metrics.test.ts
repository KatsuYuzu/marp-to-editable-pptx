const { PNG } = require('pngjs')
const {
  computeRegionalDiff,
  normalizePngPair,
} = require('./visual-diff-metrics')

describe('computeRegionalDiff', () => {
  it('detects a concentrated difference hidden by whole-slide averaging', () => {
    const width = 40
    const height = 10
    const source = new Uint8Array(width * height * 4)
    const output = new Uint8Array(source)
    source.fill(240)
    output.set(source)

    for (let y = 0; y < 5; y++) {
      for (let x = 0; x < 5; x++) {
        output[(y * width + x) * 4] = 0
      }
    }

    const result = computeRegionalDiff(source, output, width, height, {
      columns: 8,
      rows: 2,
      pixelThreshold: 0.18,
    })

    expect(result.wholeSlidePct).toBeLessThan(0.07)
    expect(result.maxRegionPct).toBeGreaterThan(0.9)
    expect(result.regions[0]).toMatchObject({
      column: 0,
      row: 0,
      width: 5,
      height: 5,
    })
  })

  it('returns zero for identical renders', () => {
    const pixels = new Uint8Array(100)
    pixels.fill(128)

    const result = computeRegionalDiff(pixels, pixels, 5, 5, {
      columns: 1,
      rows: 1,
    })

    expect(result.wholeSlidePct).toBe(0)
    expect(result.maxRegionPct).toBe(0)
  })

  it('maps the full comparison image onto the reference canvas', () => {
    const reference = new PNG({ width: 2, height: 2 })
    const comparison = new PNG({ width: 4, height: 4 })
    reference.data.set([255, 0, 0, 255], 0)
    reference.data.set([255, 0, 0, 255], 4)
    reference.data.set([0, 0, 255, 255], 8)
    reference.data.set([0, 0, 255, 255], 12)
    for (let y = 0; y < 4; y++) {
      for (let x = 0; x < 4; x++) {
        comparison.data.set(y < 2
          ? [255, 0, 0, 255]
          : [0, 0, 255, 255], (y * 4 + x) * 4)
      }
    }

    const normalized = normalizePngPair(reference, comparison)
    expect(normalized.comparison.width).toBe(reference.width)
    expect(normalized.comparison.height).toBe(reference.height)
    expect(computeRegionalDiff(
      normalized.reference.data,
      normalized.comparison.data,
      reference.width,
      reference.height,
      { columns: 1, rows: 1 },
    ).wholeSlidePct).toBe(0)
  })
})
