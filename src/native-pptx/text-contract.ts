export function normalizeContractText(text: string): string {
  return text.replace(/\s+|\ufe0f/gu, '')
}

export function findMissingContractText(
  expectedTexts: string[],
  actualText: string,
): string | undefined {
  const normalizedActual = normalizeContractText(actualText)
  let cursor = 0

  for (const expectedText of expectedTexts) {
    const normalizedExpected = normalizeContractText(expectedText)
    if (!normalizedExpected) continue
    const index = normalizedActual.indexOf(normalizedExpected, cursor)
    if (index < 0) return normalizedExpected
    cursor = index + normalizedExpected.length
  }

  return undefined
}
