import { waitWithCancellation } from './abort'

describe('waitWithCancellation', () => {
  it('rejects while an uncancellable operation is still pending', async () => {
    const controller = new AbortController()
    const operation = new Promise<never>(() => undefined)
    const waiting = waitWithCancellation(operation, controller.signal)

    controller.abort(new DOMException('cancelled', 'AbortError'))

    await expect(waiting).rejects.toMatchObject({ name: 'AbortError' })
  })
})
