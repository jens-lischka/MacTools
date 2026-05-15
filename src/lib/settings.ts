/**
 * Settings storage for the add-in.
 *
 * Two scopes, as agreed:
 *  - "document": stored inside the presentation via Office document settings;
 *    travels with the file.
 *  - "roaming": stored per-user via OfficeRuntime.storage; independent of any
 *    file, acts as the user's personal defaults.
 *
 * `get` reads the document scope first and falls back to roaming, so a
 * presentation can override a user's defaults without losing them.
 */

export type SettingsScope = "document" | "roaming";

const ROAMING_PREFIX = "mactools:";

async function readRoaming(key: string): Promise<string | null> {
  try {
    return await OfficeRuntime.storage.getItem(ROAMING_PREFIX + key);
  } catch {
    return null;
  }
}

async function writeRoaming(key: string, value: string): Promise<void> {
  await OfficeRuntime.storage.setItem(ROAMING_PREFIX + key, value);
}

function readDocument(key: string): string | null {
  const raw = Office.context.document.settings.get(key);
  return typeof raw === "string" ? raw : null;
}

function writeDocument(key: string, value: string): Promise<void> {
  Office.context.document.settings.set(key, value);
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}

/** Read a typed setting. Document scope wins over roaming when both exist. */
export async function getSetting<T>(key: string, fallback: T): Promise<T> {
  const docValue = readDocument(key);
  const raw = docValue ?? (await readRoaming(key));
  if (raw == null) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

/** Write a typed setting to the chosen scope. */
export async function setSetting<T>(
  key: string,
  value: T,
  scope: SettingsScope,
): Promise<void> {
  const raw = JSON.stringify(value);
  if (scope === "document") await writeDocument(key, raw);
  else await writeRoaming(key, raw);
}
