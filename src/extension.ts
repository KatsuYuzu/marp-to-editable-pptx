import { mkdir, unlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import type { marpCli as MarpCliFn } from '@marp-team/marp-cli'
import { nanoid } from 'nanoid'
import {
  commands,
  ExtensionContext,
  ProgressLocation,
  Uri,
  window,
} from 'vscode'
import { detectBrowserPath } from './native-pptx/browser'
import { generateNativePptx } from './native-pptx/index'

export function activate(context: ExtensionContext) {
  context.subscriptions.push(
    commands.registerCommand('marpEditablePptx.export', exportCommand),
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
      const tmpDir = tmpdir()
      const tmpId = nanoid()
      const htmlTmpPath = path.join(tmpDir, `.marp-editable-pptx-${tmpId}.html`)

      try {
        // Step 1: Convert Markdown → HTML via @marp-team/marp-cli
        const { marpCli } = (await import('@marp-team/marp-cli')) as {
          marpCli: typeof MarpCliFn
        }

        const exitCode = await marpCli(
          [doc.uri.fsPath, '-o', htmlTmpPath, '--allow-local-files'],
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
        try {
          await unlink(htmlTmpPath)
        } catch {
          // ignore
        }
      }
    },
  )
}
