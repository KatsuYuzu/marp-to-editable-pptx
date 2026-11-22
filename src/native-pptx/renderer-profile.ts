import windowsPowerPointProfile from './renderer-profiles/windows-powerpoint.json'

export type GlyphCategory = 'latin' | 'cjk'

export interface FontOverride {
  fonts: string[]
  glyphCategory: GlyphCategory
  fontSizeScale: number
}

export interface RendererProfile {
  id: string
  text: {
    cjkFontSizeScale: number
    singleLineLeadingDivisor: number
    multiLineLeadingDivisor: number
    minimumLineSpacingMultiple: number
    fontOverrides?: FontOverride[]
  }
  layout: {
    columnWidthScale: number
    nowrapWidthScale: number
    listCjkFontSizeScale: number
    badgeAverageGlyphWidthEm: number
    defaultBulletIndentPt: number
  }
  code: {
    lineSpacing: number
  }
}

export const WINDOWS_POWERPOINT_PROFILE =
  windowsPowerPointProfile as RendererProfile

export function resolveRendererProfile(
  renderer = 'windows-powerpoint',
): RendererProfile {
  if (renderer === WINDOWS_POWERPOINT_PROFILE.id) {
    return WINDOWS_POWERPOINT_PROFILE
  }
  throw new Error(`Unsupported renderer profile: ${renderer}`)
}

export function resolveFontSizeScale(
  profile: RendererProfile,
  fontFamily: string | undefined,
  glyphCategory: GlyphCategory,
): number {
  const normalizedFamily = fontFamily?.toLowerCase() ?? ''
  const override = profile.text.fontOverrides?.find(
    (candidate) =>
      candidate.glyphCategory === glyphCategory &&
      candidate.fonts.some((font) =>
        normalizedFamily.includes(font.toLowerCase()),
      ),
  )
  if (override) return override.fontSizeScale
  return glyphCategory === 'cjk' ? profile.text.cjkFontSizeScale : 1
}
