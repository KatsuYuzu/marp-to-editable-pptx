import { mkdir, unlink, writeFile } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import type { marpCli as MarpCliFn } from '@marp-team/marp-cli'
import { nanoid } from 'nanoid'
import {
  commands,
  ExtensionContext,
  ProgressLocation,
  Uri,
  window,
  workspace,
} from 'vscode'
import { detectBrowserPath } from './native-pptx/browser'
import { generateNativePptx } from './native-pptx/index'
import { buildMarpConfig, mergeMarpConfig, resolveThemeSet } from './marp-cli-config'

export function activate(context: ExtensionContext) {
  context.subscriptions.push(
    commands.registerCommand('marpToEditablePptx.export', exportCommand),
  )
}

export function deactivate() {
  // no-op
}

async function exportCommand(): Promise<void> {
  const editor = window.activeTextEditor
  if (!editor) {
    window.showErrorMessage('No active Markdown file.')
    return
  }

  const doc = editor.document
  if (doc.languageId !== 'markdown') {
    window.showErrorMessage('The active file is not a Markdown file.')
    return
  }

  if (doc.uri.scheme !== 'file') {
    window.showErrorMessage(
      'Please save the file to a local folder before exporting.',
    )
    return
  }

  if (doc.isDirty) {
    const answer = await window.showWarningMessage(
      'The file has unsaved changes. Save before exporting?',
      { modal: true },
      'Save and Export',
    )
    if (answer !== 'Save and Export') return
    await doc.save()
  }

  const defaultUri = Uri.file(doc.uri.fsPath.replace(/\.md$/i, '.pptx'))
  const saveUri = await window.showSaveDialog({
    defaultUri,
    filters: { PowerPoint: ['pptx'] },
    title: 'Export to Editable PPTX',
  })
  if (!saveUri) return

  await window.withProgress(
    {
      location: ProgressLocation.Notification,
      title: 'Exporting editable PPTX…',
      cancellable: false,
    },
    async () => {
      const tmpId = nanoid()
      // Place the temporary HTML next to the source Markdown so that
      // marp-cli resolves relative image paths (e.g. .attachments/) from
      // the correct directory. Using os.tmpdir() breaks relative paths
      // because marp-cli resolves media relative to the HTML output location.
      const htmlTmpPath = path.join(
        path.dirname(doc.uri.fsPath),
        `.marp-editable-pptx-${tmpId}.html`,
      )

      // Resolve the directory used as the project root when resolving relative
      // theme paths from the `markdown.marp.themes` setting. Prefer the
      // workspace folder that owns the document, falling back to the Markdown
      // file's own directory (matches marp-vscode's base-directory logic).
      const workspaceFolder = workspace.getWorkspaceFolder(doc.uri)
      const marpWorkingDir =
        workspaceFolder?.uri.scheme === 'file'
          ? workspaceFolder.uri.fsPath
          : path.dirname(doc.uri.fsPath)

      // Temporary files (generated config + downloaded remote themes) to remove
      // once the conversion finishes.
      const tmpFiles: string[] = []

      try {
        // Step 1: Convert Markdown → HTML via @marp-team/marp-cli
        const { marpCli } = (await import('@marp-team/marp-cli')) as {
          marpCli: typeof MarpCliFn
        }

        const marpConfig = workspace.getConfiguration('markdown.marp')
        const previewConfig = workspace.getConfiguration(
          'markdown.preview',
          doc.uri,
        )

        // Resolve custom themes from the `markdown.marp.themes` setting the same
        // way marp-vscode does, then download any remote (http/https) themes to
        // temporary files so Marp CLI can read them from disk. Remote themes are
        // only fetched in trusted workspaces. See #19.
        const { local, remote } = resolveThemeSet(
          marpConfig.get<string[]>('themes'),
          marpWorkingDir,
        )
        const themeSet = [...local]
        if (workspace.isTrusted) {
          for (const url of remote) {
            const cssPath = path.join(
              os.tmpdir(),
              `.marp-editable-pptx-theme-${nanoid()}.css`,
            )
            await writeFile(cssPath, await fetchThemeCss(url))
            tmpFiles.push(cssPath)
            themeSet.push(cssPath)
          }
        }

        // Build a Marp CLI config from the VS Code settings that drive the Marp
        // preview, so the exported PPTX matches what the user sees.
        const settingsConfig = buildMarpConfig({
          html: marpConfig.get<string>('html'),
          mathTypesetting: marpConfig.get<string>('mathTypesetting'),
          breaks: marpConfig.get<string>('breaks'),
          previewBreaks: previewConfig.get<boolean>('breaks'),
          typographer: previewConfig.get<boolean>('typographer'),
          themeSet,
        })

        // Discover a project config file (.marprc.yml, marp.config.*, a "marp"
        // key in package.json, …) exactly the way Marp CLI does — with
        // cosmiconfig, searching upward from the Markdown file's directory. It
        // is used as the base config, with the VS Code settings layered on top
        // (settings win). Only read in trusted workspaces because cosmiconfig
        // may execute JavaScript config files. See ADR-49 / issue #19.
        let fileConfig: Record<string, unknown> | undefined
        let configDir = os.tmpdir()
        if (workspace.isTrusted) {
          try {
            const { cosmiconfig } = await import('cosmiconfig')
            const found = await cosmiconfig('marp').search(
              path.dirname(doc.uri.fsPath),
            )
            if (found && !found.isEmpty && found.config) {
              fileConfig = found.config as Record<string, unknown>
              // Write the merged config next to the original file so that any
              // relative paths it contains still resolve correctly.
              configDir = path.dirname(found.filepath)
            }
          } catch {
            // Ignore unreadable/invalid project config; fall back to settings.
          }
        }

        // Merge and pass via `-c`. Using `-c` disables Marp CLI's own config
        // discovery, so merging ourselves is what lets both .marprc.yml and the
        // VS Code settings take effect together.
        const marpCliConfig = mergeMarpConfig(fileConfig, settingsConfig)

        const configPath = path.join(
          configDir,
          `.marp-editable-pptx-conf-${tmpId}.json`,
        )
        await writeFile(configPath, JSON.stringify(marpCliConfig))
        tmpFiles.push(configPath)

        const exitCode = await marpCli(
          [doc.uri.fsPath, '-o', htmlTmpPath, '-c', configPath],
          {},
        )

        if (exitCode !== 0) {
          throw new Error(`Marp CLI exited with code ${exitCode}`)
        }

        // Step 2: Detect Chromium browser
        const browserPath = detectBrowserPath('auto', undefined)
        if (!browserPath) {
          throw new Error(
            'Could not find a Chromium-based browser required for PPTX export. ' +
              'Please install Google Chrome or Microsoft Edge.',
          )
        }

        // Step 3: Generate editable PPTX from HTML
        const pptxBuffer = await generateNativePptx({
          htmlPath: htmlTmpPath,
          browserPath,
        })

        // Step 4: Write output
        await mkdir(path.dirname(saveUri.fsPath), { recursive: true })
        await writeFile(saveUri.fsPath, pptxBuffer)

        window.showInformationMessage(
          `Exported: ${path.basename(saveUri.fsPath)}`,
        )
      } finally {
        for (const file of [htmlTmpPath, ...tmpFiles]) {
          try {
            await unlink(file)
          } catch {
            // ignore
          }
        }
      }
    },
  )
}

/**
 * Download a remote theme CSS file, mirroring marp-vscode's 5-second timeout.
 */
async function fetchThemeCss(url: string): Promise<string> {
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), 5000)
  try {
    const response = await fetch(url, { signal: controller.signal })
    if (!response.ok) {
      throw new Error(`Failed fetching theme ${url} (${response.status})`)
    }
    return await response.text()
  } finally {
    clearTimeout(timeout)
  }
}
