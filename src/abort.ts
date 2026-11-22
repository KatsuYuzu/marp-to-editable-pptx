export function abortReason(signal: AbortSignal): Error | DOMException {
  const reason = signal.reason
  if (reason instanceof Error || reason instanceof DOMException) return reason
  return new DOMException('The operation was cancelled', 'AbortError')
}

export async function waitWithCancellation<T>(
  operation: Promise<T> | T,
  signal?: AbortSignal,
): Promise<T> {
  signal?.throwIfAborted()
  const pending = Promise.resolve(operation)
  if (!signal) return pending

  return await new Promise<T>((resolve, reject) => {
    const abort = () => reject(abortReason(signal))
    signal.addEventListener('abort', abort, { once: true })
    pending.then(resolve, reject).finally(() => {
      signal.removeEventListener('abort', abort)
    })
  })
}
