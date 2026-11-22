import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { PNG } from 'pngjs'
const { compareDirectories } = require('./visual-contract')

describe('compareDirectories', () => {
  it('reports concentrated regional differences and slide-count mismatches', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'visual-contract-'))
    const html = path.join(root, 'html')
    const pptx = path.join(root, 'pptx')
    fs.mkdirSync(html)
    fs.mkdirSync(pptx)

    const a = new PNG({ width: 20, height: 20 })
    const b = new PNG({ width: 20, height: 20 })
    a.data.fill(255)
    b.data.fill(255)
    for (let y = 0; y < 4; y++) {
      for (let x = 0; x < 4; x++) {
        b.data[(y * 20 + x) * 4] = 0
      }
    }
    fs.writeFileSync(path.join(html, 'html-slide-001.png'), PNG.sync.write(a))
    fs.writeFileSync(path.join(pptx, 'pptx-slide-001.png'), PNG.sync.write(b))

    const results = compareDirectories(html, pptx)
    expect(results[0].wholeSlidePct).toBeLessThan(0.2)
    expect(results[0].maxRegionPct).toBeGreaterThan(0.5)
    expect(
      compareDirectories(html, pptx, { regionFailThreshold: 0.5 })[0].status,
    ).toBe('FAIL')
    const baseline = {
      wholeSlideTolerancePct: 0.5,
      maxRegionTolerancePct: 0.5,
      slides: [{
        wholeSlidePct: results[0].wholeSlidePct * 100,
        maxRegionPct: results[0].maxRegionPct * 100,
      }],
    }
    expect(compareDirectories(html, pptx, { baseline })[0].status).toBe('OK')
    const degraded = { ...b }
    for (let y = 0; y < 10; y++) {
      for (let x = 0; x < 10; x++) {
        degraded.data[(y * 20 + x) * 4] = 0
      }
    }
    fs.writeFileSync(
      path.join(pptx, 'pptx-slide-001.png'),
      PNG.sync.write(degraded),
    )
    expect(compareDirectories(html, pptx, { baseline })[0].status).toBe('FAIL')

    fs.writeFileSync(
      path.join(html, 'html-slide-002.png'),
      PNG.sync.write(a),
    )
    expect(() => compareDirectories(html, pptx)).toThrow('Slide count mismatch')
    fs.rmSync(root, { recursive: true, force: true })
  })

  it('normalizes different render resolutions before comparing pixels', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'visual-contract-size-'))
    const html = path.join(root, 'html')
    const pptx = path.join(root, 'pptx')
    fs.mkdirSync(html)
    fs.mkdirSync(pptx)

    const reference = new PNG({ width: 16, height: 9 })
    const rendered = new PNG({ width: 80, height: 45 })
    for (let y = 0; y < reference.height; y++) {
      for (let x = 0; x < reference.width; x++) {
        const pixel = [x < 8 ? 240 : 20, y < 5 ? 30 : 210, 100, 255]
        const referenceOffset = (y * reference.width + x) * 4
        reference.data.set(pixel, referenceOffset)
      }
    }
    for (let y = 0; y < rendered.height; y++) {
      for (let x = 0; x < rendered.width; x++) {
        const pixel = [x < 40 ? 240 : 20, y < 25 ? 30 : 210, 100, 255]
        const renderedOffset = (y * rendered.width + x) * 4
        rendered.data.set(pixel, renderedOffset)
      }
    }
    fs.writeFileSync(
      path.join(html, 'html-slide-001.png'),
      PNG.sync.write(reference),
    )
    fs.writeFileSync(
      path.join(pptx, 'pptx-slide-001.png'),
      PNG.sync.write(rendered),
    )

    const [result] = compareDirectories(html, pptx)
    expect(result).toMatchObject({
      width: 16,
      height: 9,
      wholeSlidePct: 0,
      maxRegionPct: 0,
      status: 'OK',
    })
    fs.rmSync(root, { recursive: true, force: true })
  })
})
