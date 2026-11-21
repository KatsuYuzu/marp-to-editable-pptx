export interface VisualGateInput {
  wholeSlidePct: number
  maxRegionPct: number
  wholeSlideBaselinePct: number
  maxRegionBaselinePct: number
}

export interface VisualGateLimits {
  absoluteFailPct: number
  wholeSlideTolerancePct: number
  maxRegionTolerancePct: number
}

export interface VisualGateResult {
  passed: boolean
  wholeSlideLimitPct: number
  maxRegionLimitPct: number
  failures: string[]
}

export function evaluateVisualGate(
  input: VisualGateInput,
  limits: VisualGateLimits,
): VisualGateResult {
  const grandfatheredAbsolute =
    input.wholeSlideBaselinePct > limits.absoluteFailPct
  const absoluteLimit = grandfatheredAbsolute
    ? Number.POSITIVE_INFINITY
    : limits.absoluteFailPct
  const baselineLimit =
    input.wholeSlideBaselinePct + limits.wholeSlideTolerancePct
  const wholeSlideLimitPct = Math.min(absoluteLimit, baselineLimit)
  const maxRegionLimitPct =
    input.maxRegionBaselinePct + limits.maxRegionTolerancePct
  const failures: string[] = []

  if (input.wholeSlidePct > wholeSlideLimitPct) {
    failures.push(
      `whole-slide ${input.wholeSlidePct.toFixed(2)}% exceeds ` +
        `${wholeSlideLimitPct.toFixed(2)}%`,
    )
  }
  if (input.maxRegionPct > maxRegionLimitPct) {
    failures.push(
      `localized region ${input.maxRegionPct.toFixed(2)}% exceeds ` +
        `${maxRegionLimitPct.toFixed(2)}%`,
    )
  }

  return {
    passed: failures.length === 0,
    wholeSlideLimitPct,
    maxRegionLimitPct,
    failures,
  }
}
