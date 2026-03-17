import { execFile } from 'child_process';

export function openBrowser(url: string): void {
  const platform = process.platform;
  let cmd: string;
  let args: string[];

  switch (platform) {
    case 'darwin':
      cmd = 'open';
      args = [url];
      break;
    case 'win32':
      cmd = 'cmd';
      args = ['/c', 'start', '', url];
      break;
    default:
      cmd = 'xdg-open';
      args = [url];
      break;
  }

  execFile(cmd, args, (error) => {
    if (error) {
      console.error(`Warning: Could not open browser automatically (${error.message}). Please visit the URL shown above manually.`);
    }
  });
}
