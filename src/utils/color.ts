/**
 * Normalizes any standard hex color configuration input string into a standard 8-character
 * alpha-prefixed ARGB hex value required by the OpenXML Spreadsheet architecture.
 * Drops leading pound symbols (#) and manages short-hand string definitions automatically.
 *
 * @param hex Input color string (e.g., "#FFF", "8B0000", "#FF8B0000")
 * @param defaultAlpha Optional default alpha layer override value (defaults to "FF")
 */
export function normalizeToArgb(hex: string, defaultAlpha: string = "FF"): string {
  let clean = hex.replace("#", "").trim().toUpperCase();

  // Handle standard 3-character short hex formats (e.g. "FFF" -> "FFFFFF")
  if (clean.length === 3) {
    clean = clean
      .split("")
      .map((char) => char + char)
      .join("");
  }

  // If it's already a full 8-character ARGB specification block, return it directly
  if (clean.length === 8) {
    return clean;
  }

  // If it's a standard 6-character hex value, prepend the full opacity alpha flag layers
  if (clean.length === 6) {
    return defaultAlpha + clean;
  }

  // Throw descriptive structural errors for malformed formatting inputs
  throw new Error(`Invalid hexadecimal color layout length or signature encountered: "${hex}"`);
}

/**
 * Validates whether a given string matches an acceptable CSS/Spreadsheet color value format.
 */
export function isValidHexColor(hex: string): boolean {
  try {
    normalizeToArgb(hex);
    return true;
  } catch {
    return false;
  }
}
