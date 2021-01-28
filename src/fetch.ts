import _fetch from 'node-fetch'

declare global {
  const fetch: typeof _fetch
}

// @ts-ignore
if (!globalThis.fetch) {
  // @ts-ignore
  globalThis.fetch = _fetch
}

export {}
