import { mkdir, unlink, writeFile } from 'node:fs/promises'
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
import { buildMarpCliArgs } from './marp-cli-args'

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

      // Resolve the directory Marp CLI should treat as the project root.
      // Marp CLI discovers configuration files (.marprc.yml, marp.config.js,
      // a "marp" key in package.json, …) with cosmiconfig starting from
      // process.cwd(). Inside the VS Code extension host that cwd is unrelated
      // to the user's workspace, so project config — and any custom theme it
      // registers — is never loaded. Prefer the workspace folder that owns the
      // document, falling back to the Markdown file's own directory. See #19.
      const workspaceFolder = workspace.getWorkspaceFolder(doc.uri)
      const marpWorkingDir =
        workspaceFolder?.uri.scheme === 'file'
          ? workspaceFolder.uri.fsPath
          : path.dirname(doc.uri.fsPath)

      try {
        // Step 1: Convert Markdown → HTML via @marp-team/marp-cli
        const { marpCli } = (await import('@marp-team/marp-cli')) as {
          marpCli: typeof MarpCliFn
        }

        const marpConfig = workspace.getConfiguration('markdown.marp')

        // Forward --html when the user has set markdown.marp.html to 'all'.
        // This preserves <script> tags in the marp-cli HTML output, which is
        // required for runtime rendering (e.g. mermaid.js via div.mermaid).
        // Matches the same logic used by marp-vscode's marpCoreOptionForCLI.
        // Custom themes registered through `markdown.marp.themes` are forwarded
        // as `--theme-set` (see buildMarpCliArgs / issue #19).
        const marpCliArgs = buildMarpCliArgs({
          markdownPath: doc.uri.fsPath,
          htmlOutputPath: htmlTmpPath,
          workingDir: marpWorkingDir,
          htmlSetting: marpConfig.get<string>('html'),
          themes: marpConfig.get<string[]>('themes'),
        })

        // Run Marp CLI with the working directory pointed at the project root
        // so cosmiconfig can locate the configuration file. cwd is restored
        // immediately afterwards to avoid leaking global state.
        const previousCwd = process.cwd()
        let exitCode: number
        try {
          process.chdir(marpWorkingDir)
          exitCode = await marpCli(marpCliArgs, {})
        } finally {
          process.chdir(previousCwd)
        }

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
        try {
          await unlink(htmlTmpPath)
        } catch {
          // ignore
        }
      }
    },
  )
}
