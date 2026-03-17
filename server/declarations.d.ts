declare module 'officeparser' {
  export function parseOffice(
    buffer: Buffer,
    callback: (data: string, err: string | null) => void
  ): void;
}

declare module 'node-persist' {
  interface Storage {
    init(options?: { dir?: string; logging?: boolean }): Promise<void>;
    getItem(key: string): Promise<any>;
    setItem(key: string, value: any): Promise<void>;
    removeItem(key: string): Promise<void>;
  }
  const storage: Storage;
  export default storage;
}

declare module 'isomorphic-dompurify' {
  const DOMPurify: {
    sanitize(html: string, config?: Record<string, unknown>): string;
  };
  export default DOMPurify;
}
