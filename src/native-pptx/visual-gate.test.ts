import { evaluateVisualGate } from './visual-gate'

describe('evaluateVisualGate', () => {
  const limits = {
    absoluteFailPct: 6,
    wholeSlideTolerancePct: 1,
    maxRegionTolerancePct: 3,
  }

  it('rejects growth below the old absolute threshold', () => {
    expect(
      evaluateVisualGate(
        {
          wholeSlidePct: 3,
          maxRegionPct: 20,
          wholeSlideBaselinePct: 1,
          maxRegionBaselinePct: 20,
        },
        limits,
      ).passed,
    ).toBe(false)
  })

  it('rejects localized regressions', () => {
    expect(
      evaluateVisualGate(
        {
          wholeSlidePct: 1.5,
          maxRegionPct: 24,
          wholeSlideBaselinePct: 1,
          maxRegionBaselinePct: 20,
        },
        limits,
      ).passed,
    ).toBe(false)
  })

  it('grandfathers only the established absolute exceptions', () => {
    const result = evaluateVisualGate(
      {
        wholeSlidePct: 7,
        maxRegionPct: 30,
        wholeSlideBaselinePct: 6.5,
        maxRegionBaselinePct: 30,
      },
      limits,
    )

    expect(result.passed).toBe(true)
    expect(result.wholeSlideLimitPct).toBe(7.5)
  })
})
