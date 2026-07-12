import path from 'node:path'

export interface BuildMarpCliArgsParams {
  /** Absolute path to the source Markdown file. */
  markdownPath: string
  /** Absolute path where Marp CLI should write the intermediate HTML. */
  htmlOutputPath: string
  /**
   * Project root used to resolve relative theme paths. Marp CLI is also run
   * with this directory as its working directory so that configuration files
   * (`.marprc.yml`, `marp.config.*`, …) are discovered. See issue #19.
   */
  workingDir: string
  /** Value of the `markdown.marp.html` setting, if any. */
  htmlSetting?: string
  /** Value of the `markdown.marp.themes` setting, if any. */
  themes?: readonly string[]
}

/**
 * Build the argument vector for `@marp-team/marp-cli` when converting a Marp
 * Markdown document to HTML.
 *
 * Custom themes registered through the `markdown.marp.themes` setting are
 * forwarded as `--theme-set` so the exported PPTX matches the Marp preview.
 * Relative theme paths are resolved against {@link BuildMarpCliArgsParams.workingDir}.
 * Blank entries are ignored and remote (`http`/`https`) theme URLs are left to
 * Marp CLI's own configuration resolution.
 */
export function buildMarpCliArgs({
  markdownPath,
  htmlOutputPath,
  workingDir,
  htmlSetting,
  themes,
}: BuildMarpCliArgsParams): string[] {
  const themeSetArgs = (themes ?? [])
    .filter((theme) => typeof theme === 'string' && theme.trim() !== '')
    .filter((theme) => !/^https?:\/\//i.test(theme))
    .flatMap((theme) => ['--theme-set', path.resolve(workingDir, theme)])

  return [
    markdownPath,
    '-o',
    htmlOutputPath,
    '--allow-local-files',
    ...(htmlSetting === 'all' ? ['--html'] : []),
    ...themeSetArgs,
  ]
}
