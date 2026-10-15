import { existsSync } from 'node:fs'
import { homedir } from 'node:os'
import path from 'node:path'
import {
  Browser,
  BrowserPlatform,
  ChromeReleaseChannel,
  computeSystemExecutablePath,
} from '@puppeteer/browsers'

/**
 * Detect a usable Chromium-based browser executable.
 *
 * Uses Puppeteer's system Chrome lookup and standard Edge installation paths.
 *
 * @param preference - User-configured browser preference ('auto' | 'chrome' | 'edge')
 * @param explicitPath - An explicit path to a browser binary (overrides detection)
 */
export function detectBrowserPath(
  preference = 'auto',
  explicitPath?: string,
): string | undefined {
  if (explicitPath && existsSync(explicitPath)) return explicitPath

  const platform = detectPlatform()
  if (!platform) return undefined

  const candidates: { browser: Browser; channel: ChromeReleaseChannel }[] = []

  if (preference === 'chrome' || preference === 'auto') {
    candidates.push({
      browser: Browser.CHROME,
      channel: ChromeReleaseChannel.STABLE,
    })
  }

  for (const { browser, channel } of candidates) {
    try {
      const executablePath = computeSystemExecutablePath({
        browser,
        platform,
        channel,
      })
      if (existsSync(executablePath)) return executablePath
    } catch {
      // Not found — try next candidate
    }
  }

  if (preference === 'edge' || preference === 'auto') {
    const edgePath = edgeExecutablePaths().find((candidate) =>
      existsSync(candidate),
    )
    if (edgePath) return edgePath
  }

  return undefined
}

function edgeExecutablePaths(): string[] {
  switch (process.platform) {
    case 'win32':
      return [
        process.env['PROGRAMFILES(X86)'],
        process.env.PROGRAMFILES,
        process.env.LOCALAPPDATA,
      ]
        .filter((root): root is string => !!root)
        .map((root) =>
          path.win32.join(root, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
        )
    case 'darwin':
      return [
        '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
        path.posix.join(
          homedir(),
          'Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
        ),
      ]
    case 'linux':
      return ['/usr/bin/microsoft-edge', '/usr/bin/microsoft-edge-stable']
    default:
      return []
  }
}

function detectPlatform(): BrowserPlatform | undefined {
  switch (process.platform) {
    case 'win32':
      return process.arch === 'x64'
        ? BrowserPlatform.WIN64
        : BrowserPlatform.WIN32
    case 'darwin':
      return BrowserPlatform.MAC
    case 'linux':
      return BrowserPlatform.LINUX
    default:
      return undefined
  }
}
