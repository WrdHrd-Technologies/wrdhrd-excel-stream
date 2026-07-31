export function escapeXml(unsafe: string | number | boolean): string {
  const str = String(unsafe);
  let result = "";
  let lastIdx = 0;
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    let replacement: string | null = null;
    switch (code) {
      case 38:
        replacement = "&amp;";
        break; // &
      case 60:
        replacement = "&lt;";
        break; // <
      case 62:
        replacement = "&gt;";
        break; // >
      case 34:
        replacement = "&quot;";
        break; // "
      case 39:
        replacement = "&apos;";
        break; // '
    }
    if (replacement !== null) {
      result += str.substring(lastIdx, i) + replacement;
      lastIdx = i + 1;
    }
  }
  return result + str.substring(lastIdx);
}
