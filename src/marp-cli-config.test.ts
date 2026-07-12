import path from 'node:path'
import { buildMarpConfig, resolveThemeSet } from './marp-cli-config'

const ROOT = path.resolve(path.sep, 'workspace', 'root')

describe('resolveThemeSet (issue #19 / marp-vscode alignment)', () => {
  it('returns empty sets for undefined or empty input', () => {
    expect(resolveThemeSet(undefined, ROOT)).toEqual({ local: [], remote: [] })
    expect(resolveThemeSet([], ROOT)).toEqual({ local: [], remote: [] })
  })

  it('resolves a relative local theme against the root directory', () => {
    const { local, remote } = resolveThemeSet(['./themes/a.css'], ROOT)
    expect(local).toEqual([path.resolve(ROOT, 'themes/a.css')])
    expect(remote).toEqual([])
  })

  it('classifies http(s) URLs as remote and keeps them verbatim', () => {
    const { local, remote } = resolveThemeSet(
      ['https://example.com/a.css', 'http://example.com/b.css'],
      ROOT,
    )
    expect(local).toEqual([])
    expect(remote).toEqual([
      'https://example.com/a.css',
      'http://example.com/b.css',
    ])
  })

  it('drops entries that escape the root directory (traversal guard)', () => {
    const { local } = resolveThemeSet(
      ['../outside.css', './inside.css'],
      ROOT,
    )
    expect(local).toEqual([path.resolve(ROOT, 'inside.css')])
  })

  it('ignores blank and non-string entries', () => {
    const { local, remote } = resolveThemeSet(
      ['', '   ', undefined as unknown as string],
      ROOT,
    )
    expect(local).toEqual([])
    expect(remote).toEqual([])
  })

  it('removes duplicates while preserving order', () => {
    const { local, remote } = resolveThemeSet(
      ['./a.css', './a.css', 'https://x/y.css', 'https://x/y.css'],
      ROOT,
    )
    expect(local).toEqual([path.resolve(ROOT, 'a.css')])
    expect(remote).toEqual(['https://x/y.css'])
  })
})

describe('buildMarpConfig (marp-vscode alignment)', () => {
  it('always enables allowLocalFiles', () => {
    expect(buildMarpConfig({}).allowLocalFiles).toBe(true)
  })

  it('omits optional sections when nothing is configured', () => {
    const config = buildMarpConfig({})
    expect(config).toEqual({ allowLocalFiles: true })
  })

  it("maps markdown.marp.html 'all' → true and 'off' → false", () => {
    expect(buildMarpConfig({ html: 'all' }).html).toBe(true)
    expect(buildMarpConfig({ html: 'off' }).html).toBe(false)
    expect(buildMarpConfig({ html: undefined }).html).toBeUndefined()
  })

  it('maps mathTypesetting to the Marp math option', () => {
    expect(buildMarpConfig({ mathTypesetting: 'off' }).options?.math).toBe(false)
    expect(buildMarpConfig({ mathTypesetting: 'katex' }).options?.math).toBe(
      'katex',
    )
    expect(buildMarpConfig({ mathTypesetting: 'mathjax' }).options?.math).toBe(
      'mathjax',
    )
    expect(buildMarpConfig({ mathTypesetting: undefined }).options).toBeUndefined()
  })

  it('maps breaks on/off directly', () => {
    expect(buildMarpConfig({ breaks: 'on' }).options?.markdown?.breaks).toBe(
      true,
    )
    expect(buildMarpConfig({ breaks: 'off' }).options?.markdown?.breaks).toBe(
      false,
    )
  })

  it("resolves breaks 'inherit' from previewBreaks", () => {
    expect(
      buildMarpConfig({ breaks: 'inherit', previewBreaks: true }).options
        ?.markdown?.breaks,
    ).toBe(true)
    expect(
      buildMarpConfig({ breaks: 'inherit', previewBreaks: false }).options
        ?.markdown?.breaks,
    ).toBe(false)
    // No preview value → nothing forwarded.
    expect(
      buildMarpConfig({ breaks: 'inherit' }).options,
    ).toBeUndefined()
  })

  it('forwards the typographer setting', () => {
    expect(
      buildMarpConfig({ typographer: true }).options?.markdown?.typographer,
    ).toBe(true)
    expect(
      buildMarpConfig({ typographer: false }).options?.markdown?.typographer,
    ).toBe(false)
  })

  it('includes a non-empty themeSet and omits an empty one', () => {
    const abs = path.resolve(ROOT, 'a.css')
    expect(buildMarpConfig({ themeSet: [abs] }).themeSet).toEqual([abs])
    expect(buildMarpConfig({ themeSet: [] }).themeSet).toBeUndefined()
  })

  it('combines html, math, breaks, typographer and themeSet', () => {
    const abs = path.resolve(ROOT, 'a.css')
    const config = buildMarpConfig({
      html: 'all',
      mathTypesetting: 'katex',
      breaks: 'off',
      typographer: true,
      themeSet: [abs],
    })
    expect(config).toEqual({
      allowLocalFiles: true,
      html: true,
      themeSet: [abs],
      options: {
        markdown: { breaks: false, typographer: true },
        math: 'katex',
      },
    })
  })
})
