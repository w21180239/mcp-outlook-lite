export function debug(...args: unknown[]): void {
  if (process.env.DEBUG) {
    console.error(...args);
  }
}

export function warn(...args: unknown[]): void {
  console.error('[WARN]', ...args);
}
