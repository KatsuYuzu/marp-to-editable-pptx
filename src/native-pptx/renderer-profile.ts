export interface RendererProfile {
  id: string
  text: {
    cjkFontSizeScale: number
    singleLineLeadingDivisor: number
    multiLineLeadingDivisor: number
    minimumLineSpacingMultiple: number
  }
  layout: {
    columnWidthScale: number
    nowrapWidthScale: number
    badgeAverageGlyphWidthEm: number
    defaultBulletIndentPt: number
  }
  code: {
    lineSpacing: number
  }
}

export const WINDOWS_POWERPOINT_PROFILE: RendererProfile = Object.freeze({
  id: 'windows-powerpoint',
  text: {
    cjkFontSizeScale: 0.98,
    singleLineLeadingDivisor: 1.15,
    multiLineLeadingDivisor: 1.2,
    minimumLineSpacingMultiple: 0.85,
  },
  layout: {
    columnWidthScale: 1.05,
    nowrapWidthScale: 1.15,
    badgeAverageGlyphWidthEm: 0.58,
    defaultBulletIndentPt: 27,
  },
  code: {
    lineSpacing: 1.2,
  },
})

export function resolveRendererProfile(
  renderer = 'windows-powerpoint',
): RendererProfile {
  if (renderer === WINDOWS_POWERPOINT_PROFILE.id) {
    return WINDOWS_POWERPOINT_PROFILE
  }
  throw new Error(`Unsupported renderer profile: ${renderer}`)
}
