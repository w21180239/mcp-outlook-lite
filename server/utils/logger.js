export function debug(...args) {
  if (process.env.DEBUG) {
    console.error(...args);
  }
}

export function warn(...args) {
  console.error('[WARN]', ...args);
}
