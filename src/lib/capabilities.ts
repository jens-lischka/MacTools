/**
 * Runtime capability detection.
 *
 * The manifest declares a low `MinVersion` so the add-in installs on every
 * client. Individual features then check at runtime whether the API set they
 * need is supported, and disable themselves cleanly when it is not — instead
 * of failing when invoked.
 */

const KNOWN_SETS = [
  "1.1", "1.2", "1.3", "1.4", "1.5",
  "1.6", "1.7", "1.8", "1.9", "1.10",
];

/** True when the given PowerPoint API requirement set is available. */
export function isApiSupported(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", version);
  } catch {
    return false;
  }
}

/** The highest PowerPoint API requirement set this client supports. */
export function highestSupportedApi(): string {
  let highest = "1.1";
  for (const version of KNOWN_SETS) {
    if (isApiSupported(version)) highest = version;
  }
  return highest;
}
