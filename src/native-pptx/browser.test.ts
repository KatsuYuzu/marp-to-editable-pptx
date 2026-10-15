import { existsSync } from 'node:fs'
import os from 'node:os'
import { detectBrowserPath } from './browser'

jest.mock('node:fs', () => ({
  existsSync: jest.fn(),
}))

jest.mock('@puppeteer/browsers', () => ({
  Browser: { CHROME: 'chrome' },
  BrowserPlatform: {
    WIN64: 'win64',
    WIN32: 'win32',
    MAC: 'mac',
    LINUX: 'linux',
  },
  ChromeReleaseChannel: { STABLE: 'stable' },
  computeSystemExecutablePath: jest.fn(),
}))

const mockedExistsSync = existsSync as jest.MockedFunction<typeof existsSync>
const originalPlatform = Object.getOwnPropertyDescriptor(process, 'platform')!
const { computeSystemExecutablePath: mockedCompute } = jest.requireMock(
  '@puppeteer/browsers',
) as {
  computeSystemExecutablePath: jest.Mock
}

describe('detectBrowserPath', () => {
  beforeEach(() => {
    jest.resetAllMocks()
    mockedExistsSync.mockReturnValue(false)
    mockedCompute.mockImplementation(() => {
      throw new Error('Not found')
    })
  })

  afterEach(() => {
    jest.restoreAllMocks()
    Object.defineProperty(process, 'platform', originalPlatform)
  })

  it('returns explicit path when it exists', () => {
    mockedExistsSync.mockReturnValue(true)
    const result = detectBrowserPath('auto', '/path/to/chrome')
    expect(result).toBe('/path/to/chrome')
  })

  it('falls back to auto-detection when explicit path does not exist', () => {
    mockedExistsSync.mockImplementation((p) => {
      if (p === '/path/to/missing') return false
      return true
    })
    mockedCompute.mockReturnValue('/detected/chrome')

    const result = detectBrowserPath('auto', '/path/to/missing')
    expect(result).toBe('/detected/chrome')
  })

  it('returns path from computeSystemExecutablePath', () => {
    mockedExistsSync.mockReturnValue(true)
    mockedCompute.mockReturnValue('/usr/bin/google-chrome')

    const result = detectBrowserPath('chrome')
    expect(result).toBe('/usr/bin/google-chrome')
    expect(mockedCompute).toHaveBeenCalled()
  })

  it('returns undefined when no browser is found', () => {
    mockedExistsSync.mockReturnValue(false)
    mockedCompute.mockReturnValue('/nonexistent/chrome')

    const result = detectBrowserPath('chrome')
    expect(result).toBeUndefined()
  })

  it('returns undefined when computeSystemExecutablePath throws', () => {
    mockedExistsSync.mockReturnValue(false)
    mockedCompute.mockImplementation(() => {
      throw new Error('Not found')
    })

    const result = detectBrowserPath('chrome')
    expect(result).toBeUndefined()
  })

  it.each([
    ['PROGRAMFILES(X86)', 'C:\\Program Files (x86)'],
    ['PROGRAMFILES', 'C:\\Program Files'],
    ['LOCALAPPDATA', 'C:\\Users\\test\\AppData\\Local'],
  ])('detects Windows Edge installed under %s', (variable, root) => {
    Object.defineProperty(process, 'platform', { value: 'win32' })
    jest.replaceProperty(process, 'env', { [variable]: root })
    const executable = `${root}\\Microsoft\\Edge\\Application\\msedge.exe`
    mockedExistsSync.mockImplementation((candidate) => candidate === executable)

    expect(detectBrowserPath('edge')).toBe(executable)
    expect(mockedCompute).not.toHaveBeenCalled()
  })

  it('detects Edge in the macOS Applications directory', () => {
    Object.defineProperty(process, 'platform', { value: 'darwin' })
    const executable =
      '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge'
    mockedExistsSync.mockImplementation((candidate) => candidate === executable)

    expect(detectBrowserPath('edge')).toBe(executable)
    expect(mockedCompute).not.toHaveBeenCalled()
  })

  it('detects Edge installed in the user Applications directory on macOS', () => {
    Object.defineProperty(process, 'platform', { value: 'darwin' })
    jest.spyOn(os, 'homedir').mockReturnValue('/Users/test')
    const executable =
      '/Users/test/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge'
    mockedExistsSync.mockImplementation((candidate) => candidate === executable)

    expect(detectBrowserPath('edge')).toBe(executable)
  })

  it.each(['/usr/bin/microsoft-edge', '/usr/bin/microsoft-edge-stable'])(
    'detects Linux Edge at %s',
    (executable) => {
      Object.defineProperty(process, 'platform', { value: 'linux' })
      mockedExistsSync.mockImplementation((candidate) => candidate === executable)

      expect(detectBrowserPath('edge')).toBe(executable)
    },
  )

  it('falls back to Edge when automatic Chrome detection fails', () => {
    Object.defineProperty(process, 'platform', { value: 'darwin' })
    const executable =
      '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge'
    mockedExistsSync.mockImplementation((candidate) => candidate === executable)

    expect(detectBrowserPath()).toBe(executable)
    expect(mockedCompute).toHaveBeenCalled()
  })

  it('does not fall back to Edge when Chrome is explicitly selected', () => {
    Object.defineProperty(process, 'platform', { value: 'darwin' })
    mockedExistsSync.mockReturnValue(true)

    expect(detectBrowserPath('chrome')).toBeUndefined()
  })

  it('prefers Chrome over Edge in auto mode when both exist', () => {
    mockedCompute.mockReturnValue('/detected/chrome')
    mockedExistsSync.mockReturnValue(true)

    expect(detectBrowserPath()).toBe('/detected/chrome')
  })

  it('uses an explicit path even when a different browser is selected', () => {
    mockedExistsSync.mockReturnValue(true)

    expect(detectBrowserPath('edge', '/custom/chrome')).toBe('/custom/chrome')
    expect(mockedCompute).not.toHaveBeenCalled()
  })

  it('returns undefined when Edge is selected but not installed', () => {
    expect(detectBrowserPath('edge')).toBeUndefined()
    expect(mockedCompute).not.toHaveBeenCalled()
  })
})
