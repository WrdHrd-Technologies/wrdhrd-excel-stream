const ESCAPE_MAP: Record<string, string> = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&apos;",
};

export function escapeXml(value: string | number | boolean | null | undefined): string {
  if (value === null || value === undefined) return "";
  const str = String(value);
  if (!/[&<>"']/.test(str)) {
    return str;
  }
  return str.replace(/[&<>"']/g, (char) => ESCAPE_MAP[char] || char);
}
