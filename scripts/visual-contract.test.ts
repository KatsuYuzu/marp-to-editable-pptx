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

    fs.writeFileSync(
      path.join(html, 'html-slide-002.png'),
      PNG.sync.write(a),
    )
    expect(() => compareDirectories(html, pptx)).toThrow('Slide count mismatch')
    fs.rmSync(root, { recursive: true, force: true })
  })
})
