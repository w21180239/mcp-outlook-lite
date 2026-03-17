import { exec } from 'child_process';

export function openBrowser(url) {
  const platform = process.platform;
  let command;

  switch (platform) {
    case 'darwin':
      command = `open "${url}"`;
      break;
    case 'win32':
      command = `start "" "${url}"`;
      break;
    default:
      command = `xdg-open "${url}"`;
      break;
  }

  exec(command, (error) => {
    if (error) {
      // Silent fail - URL is already displayed above
    }
  });
}
