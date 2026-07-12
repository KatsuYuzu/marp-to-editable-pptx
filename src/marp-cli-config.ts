import path from 'node:path'

/**
 * Theme entries from the `markdown.marp.themes` setting, split by kind and
 * resolved the same way marp-vscode resolves them.
 */
export interface ResolvedThemeSet {
  /** Absolute file-system paths of local theme CSS files. */
  local: string[]
  /** Remote (`http`/`https`) theme URLs to download before use. */
  remote: string[]
}

const isRemote = (entry: string): boolean => /^https?:\/\//i.test(entry)

/**
 * Resolve `markdown.marp.themes` entries the way marp-vscode does:
 * - remote URLs (`http`/`https`) are kept as-is (to be downloaded),
 * - local paths are resolved against `rootDir`,
 * - entries escaping `rootDir` (directory traversal) are dropped,
 * - blank entries are ignored and duplicates removed (order preserved).
 *
 * @param rootDir Workspace folder that owns the document, or the Markdown
 *   file's directory when there is no workspace folder.
 */
export function resolveThemeSet(
  themes: readonly string[] | undefined,
  rootDir: string,
): ResolvedThemeSet {
  const local: string[] = []
  const remote: string[] = []
  const seen = new Set<string>()
  const root = path.resolve(rootDir)

  for (const entry of themes ?? []) {
    if (typeof entry !== 'string' || entry.trim() === '') continue

    if (isRemote(entry)) {
      if (!seen.has(entry)) {
        seen.add(entry)
        remote.push(entry)
      }
      continue
    }

    const resolved = path.resolve(root, entry)
    // Prevent directory traversal outside the project root.
    if (resolved !== root && !resolved.startsWith(root + path.sep)) continue

    if (!seen.has(resolved)) {
      seen.add(resolved)
      local.push(resolved)
    }
  }

  return { local, remote }
}

/** Inputs read from VS Code settings, mirroring marp-vscode. */
export interface MarpConfigInput {
  /** `markdown.marp.html` (`'all'` | `'off'` | …). */
  html?: string
  /** `markdown.marp.mathTypesetting` (`'off'` | `'katex'` | `'mathjax'`). */
  mathTypesetting?: string
  /** `markdown.marp.breaks` (`'on'` | `'off'` | `'inherit'`). */
  breaks?: string
  /** `markdown.preview.breaks`, used when `breaks` is `'inherit'`. */
  previewBreaks?: boolean
  /** `markdown.preview.typographer`. */
  typographer?: boolean
  /** Absolute local theme paths (remote entries already downloaded). */
  themeSet?: readonly string[]
}

/** A Marp CLI configuration object (subset used by this extension). */
export interface MarpConfig {
  allowLocalFiles: true
  html?: boolean
  themeSet?: string[]
  options?: {
    markdown?: { breaks?: boolean; typographer?: boolean }
    math?: 'katex' | 'mathjax' | false
  }
}

/**
 * Build a Marp CLI configuration object from VS Code settings, mirroring
 * marp-vscode's `marpCoreOptionForCLI`. Passing this to Marp CLI via `-c`
 * makes the exported PPTX match the VS Code Marp preview and, as a side
 * effect, disables discovery of ambient config files (`.marprc.yml`).
 *
 * Only settings the caller actually provides are forwarded; unset values are
 * omitted so Marp CLI keeps its own defaults.
 */
export function buildMarpConfig(input: MarpConfigInput): MarpConfig {
  const config: MarpConfig = { allowLocalFiles: true }

  if (input.html === 'all') config.html = true
  else if (input.html === 'off') config.html = false

  if (input.themeSet && input.themeSet.length > 0) {
    config.themeSet = [...input.themeSet]
  }

  const markdown: { breaks?: boolean; typographer?: boolean } = {}
  if (input.breaks === 'on') markdown.breaks = true
  else if (input.breaks === 'off') markdown.breaks = false
  else if (input.breaks === 'inherit' && input.previewBreaks !== undefined) {
    markdown.breaks = input.previewBreaks
  }
  if (input.typographer !== undefined) markdown.typographer = input.typographer

  let math: 'katex' | 'mathjax' | false | undefined
  if (input.mathTypesetting === 'off') math = false
  else if (input.mathTypesetting === 'katex') math = 'katex'
  else if (input.mathTypesetting === 'mathjax') math = 'mathjax'

  const options: NonNullable<MarpConfig['options']> = {}
  if (Object.keys(markdown).length > 0) options.markdown = markdown
  if (math !== undefined) options.math = math
  if (Object.keys(options).length > 0) config.options = options

  return config
}
