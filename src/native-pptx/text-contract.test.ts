import { findMissingContractText } from './text-contract'

describe('findMissingContractText', () => {
  it('detects a missing duplicate text occurrence', () => {
    expect(
      findMissingContractText(['Label-A', 'Label-A'], 'Label-A'),
    ).toBe('Label-A')
  })

  it('detects reordered text', () => {
    expect(
      findMissingContractText(['Label-A', 'Label-B'], 'Label-BLabel-A'),
    ).toBe('Label-B')
  })

  it('preserves single-character text', () => {
    expect(findMissingContractText(['A'], '')).toBe('A')
    expect(findMissingContractText(['A'], 'A')).toBeUndefined()
  })
})
