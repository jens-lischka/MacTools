/**
 * Minimal ambient declaration for the OfficeRuntime.storage API used by the
 * settings store. The installed @types/office-js does not declare it.
 */
declare namespace OfficeRuntime {
  const storage: {
    getItem(key: string): Promise<string | null>;
    setItem(key: string, value: string): Promise<void>;
    removeItem(key: string): Promise<void>;
    getItems(keys: string[]): Promise<Record<string, string>>;
    setItems(items: Record<string, string>): Promise<void>;
  };
}
