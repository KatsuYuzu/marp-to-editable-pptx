import path from 'node:path'
import { buildMarpCliArgs } from './marp-cli-args'

const WORKING_DIR = path.join(path.sep, 'workspace', 'root')
const MD = path.join(WORKING_DIR, 'deck.md')
const HTML = path.join(WORKING_DIR, '.marp-editable-pptx-abc.html')

const base = {
  markdownPath: MD,
  htmlOutputPath: HTML,
  workingDir: WORKING_DIR,
}

describe('buildMarpCliArgs (issue #19)', () => {
  it('produces the core conversion args by default', () => {
    expect(buildMarpCliArgs(base)).toEqual([
      MD,
      '-o',
      HTML,
      '--allow-local-files',
    ])
  })

  it("appends --html only when markdown.marp.html is 'all'", () => {
    expect(buildMarpCliArgs({ ...base, htmlSetting: 'all' })).toContain('--html')
    expect(buildMarpCliArgs({ ...base, htmlSetting: 'off' })).not.toContain(
      '--html',
    )
    expect(buildMarpCliArgs(base)).not.toContain('--html')
  })

  it('forwards a relative theme as --theme-set resolved against workingDir', () => {
    const args = buildMarpCliArgs({
      ...base,
      themes: ['./themes/my_theme.css'],
    })
    expect(args).toContain('--theme-set')
    expect(args).toContain(
      path.resolve(WORKING_DIR, './themes/my_theme.css'),
    )
  })

  it('forwards multiple themes as repeated --theme-set pairs in order', () => {
    const args = buildMarpCliArgs({
      ...base,
      themes: ['./a.css', './nested/b.css'],
    })
    expect(args.slice(-4)).toEqual([
      '--theme-set',
      path.resolve(WORKING_DIR, './a.css'),
      '--theme-set',
      path.resolve(WORKING_DIR, './nested/b.css'),
    ])
  })

  it('keeps absolute theme paths absolute', () => {
    const abs = path.join(path.sep, 'elsewhere', 'theme.css')
    const args = buildMarpCliArgs({ ...base, themes: [abs] })
    expect(args).toContain(path.resolve(abs))
  })

  it('ignores blank and whitespace-only theme entries', () => {
    const args = buildMarpCliArgs({ ...base, themes: ['', '   '] })
    expect(args).not.toContain('--theme-set')
  })

  it('leaves remote http(s) theme URLs to Marp CLI config resolution', () => {
    const args = buildMarpCliArgs({
      ...base,
      themes: ['https://example.com/t.css', 'http://example.com/u.css'],
    })
    expect(args).not.toContain('--theme-set')
  })

  it('forwards local themes while skipping remote ones', () => {
    const args = buildMarpCliArgs({
      ...base,
      themes: ['https://example.com/remote.css', './local.css'],
    })
    expect(args.filter((a) => a === '--theme-set')).toHaveLength(1)
    expect(args).toContain(path.resolve(WORKING_DIR, './local.css'))
  })

  it('combines --html and --theme-set', () => {
    const args = buildMarpCliArgs({
      ...base,
      htmlSetting: 'all',
      themes: ['./local.css'],
    })
    expect(args).toContain('--html')
    expect(args).toContain('--theme-set')
    // The Markdown path stays first and the output flag directly follows it.
    expect(args[0]).toBe(MD)
    expect(args[1]).toBe('-o')
    expect(args[2]).toBe(HTML)
  })
})
