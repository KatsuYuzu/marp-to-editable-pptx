import {
  resolveFontSizeScale,
  resolveRendererProfile,
} from './renderer-profile'

describe('resolveRendererProfile', () => {
  it('selects the Windows PowerPoint profile', () => {
    expect(resolveRendererProfile().id).toBe('windows-powerpoint')
  })

  it('uses font and glyph-specific overrides for CJK text', () => {
    const profile = resolveRendererProfile()

    expect(resolveFontSizeScale(profile, 'Segoe UI', 'cjk')).toBe(0.98)
    expect(resolveFontSizeScale(profile, 'Unknown Font', 'cjk')).toBe(0.98)
    expect(resolveFontSizeScale(profile, 'Segoe UI', 'latin')).toBe(1)
  })

  it('rejects unsupported renderer profiles', () => {
    expect(() => resolveRendererProfile('unsupported')).toThrow(
      'Unsupported renderer profile',
    )
  })
})
