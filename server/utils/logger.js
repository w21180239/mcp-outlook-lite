export function debug(...args) {
  if (process.env.DEBUG) {
    console.error(...args);
  }
}
