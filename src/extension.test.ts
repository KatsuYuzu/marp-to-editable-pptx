import { mkdir, unlink, writeFile } from 'node:fs/promises'
import { marpCli } from '@marp-team/marp-cli'
import { commands, ExtensionContext, Uri, window, workspace } from 'vscode'
import { activate } from './extension'
import { detectBrowserPath } from './native-pptx/browser'
import { generateNativePptx } from './native-pptx/index'

jest.mock('node:fs/promises', () => ({
  mkdir: jest.fn(),
  unlink: jest.fn(),
  writeFile: jest.fn(),
}))
jest.mock('nanoid', () => ({ nanoid: () => 'test-id' }))
jest.mock('@marp-team/marp-cli', () => ({ marpCli: jest.fn() }))
jest.mock('./native-pptx/browser', () => ({ detectBrowserPath: jest.fn() }))
jest.mock('./native-pptx/index', () => ({ generateNativePptx: jest.fn() }))
jest.mock(
  'vscode',
  () => ({
    commands: { registerCommand: jest.fn() },
    ProgressLocation: { Notification: 15 },
    Uri: { file: (fsPath: string) => ({ scheme: 'file', fsPath }) },
    window: {
      activeTextEditor: undefined,
      showSaveDialog: jest.fn(),
      withProgress: jest.fn(),
      showErrorMessage: jest.fn(),
      showInformationMessage: jest.fn(),
    },
    workspace: {
      isTrusted: false,
      getWorkspaceFolder: jest.fn(),
      getConfiguration: jest.fn(),
    },
  }),
  { virtual: true },
)

const documentUri = Uri.file('C:\\slides\\deck.md')
const config = {
  get: jest.fn(),
  has: jest.fn(),
  inspect: jest.fn(),
  update: jest.fn(),
}
let exportCommand: () => Promise<void>

describe('export command browser selection', () => {
  beforeEach(() => {
    jest.resetAllMocks()
    Object.assign(window, {
      activeTextEditor: {
        document: { languageId: 'markdown', uri: documentUri, isDirty: false },
      },
    })
    jest.mocked(window.showSaveDialog).mockResolvedValue(
      Uri.file('C:\\slides\\deck.pptx'),
    )
    jest.mocked(window.withProgress).mockImplementation(async (_, task) =>
      task(
        { report: jest.fn() },
        {
          isCancellationRequested: false,
          onCancellationRequested: jest.fn(),
        },
      ),
    )
    config.get.mockImplementation((_key: string, fallback?: unknown) => fallback)
    jest.mocked(workspace.getConfiguration).mockReturnValue(config)
    jest.mocked(detectBrowserPath).mockReturnValue('C:\\browser\\msedge.exe')
    jest.mocked(marpCli).mockResolvedValue(0)
    jest.mocked(generateNativePptx).mockResolvedValue(Buffer.from('pptx'))
    const context = { subscriptions: [] } as Pick<ExtensionContext, 'subscriptions'>
    activate(context as ExtensionContext)
    exportCommand = jest.mocked(commands.registerCommand).mock.calls[0][1]
  })

  it('passes document-scoped Edge settings to detection and the selected path to PPTX generation', async () => {
    const executable = 'C:\\custom\\msedge.exe'
    config.get.mockImplementation((key: string, fallback?: unknown) => {
      if (key === 'browser') return 'edge'
      if (key === 'browserPath') return executable
      return fallback
    })
    jest.mocked(detectBrowserPath).mockReturnValue(executable)

    await exportCommand()

    expect(workspace.getConfiguration).toHaveBeenCalledWith(
      'markdown.marp',
      documentUri,
    )
    expect(detectBrowserPath).toHaveBeenCalledWith('edge', executable)
    expect(generateNativePptx).toHaveBeenCalledWith({
      htmlPath: expect.stringContaining('.marp-editable-pptx-test-id.html'),
      browserPath: executable,
    })
    expect(marpCli).toHaveBeenCalledTimes(1)
    expect(mkdir).toHaveBeenCalled()
    expect(writeFile).toHaveBeenCalledWith(
      'C:\\slides\\deck.pptx',
      Buffer.from('pptx'),
    )
    expect(window.showInformationMessage).toHaveBeenCalled()
    expect(window.showErrorMessage).not.toHaveBeenCalled()
    expect(unlink).toHaveBeenCalled()
  })

  it('uses automatic detection when browser settings are unset', async () => {
    await exportCommand()

    expect(detectBrowserPath).toHaveBeenCalledWith('auto', undefined)
    expect(generateNativePptx).toHaveBeenCalled()
  })

  it('reports a missing browser with configuration guidance and does not convert or write output', async () => {
    jest.mocked(detectBrowserPath).mockReturnValue(undefined)

    await exportCommand()

    expect(window.showErrorMessage).toHaveBeenCalledWith(
      expect.stringContaining('markdown.marp.browserPath'),
    )
    expect(marpCli).not.toHaveBeenCalled()
    expect(generateNativePptx).not.toHaveBeenCalled()
    expect(writeFile).not.toHaveBeenCalled()
    expect(window.showInformationMessage).not.toHaveBeenCalled()
  })

  it('reports browser launch failures without announcing success', async () => {
    jest.mocked(generateNativePptx).mockRejectedValue(
      new Error('Browser launch failed'),
    )

    await exportCommand()

    expect(window.showErrorMessage).toHaveBeenCalledWith(
      'PPTX export failed: Browser launch failed',
    )
    expect(window.showInformationMessage).not.toHaveBeenCalled()
    expect(unlink).toHaveBeenCalledTimes(2)
  })
})
