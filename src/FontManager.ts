import { FontStyle } from "./models/CellStyle";

export class FontManager {
  private fonts: string[] = [];
  private cache = new Map<string, number>();

  constructor() {
    this.getFontIndex({ name: "Calibri", fontSize: 11 });
  }

  public getFontIndex(font?: FontStyle): number {
    if (!font) return 0;
    const key = JSON.stringify(font);
    if (this.cache.has(key)) return this.cache.get(key)!;

    let xml = "<font>";
    if (font.bold) xml += "<b/>";
    if (font.italic) xml += "<i/>";
    xml += `<sz val="${font.fontSize || 11}"/>`;
    if (font.color) xml += `<color rgb="${font.color.replace("#", "")}"/>`;
    xml += `<name val="${font.name || "Calibri"}"/></font>`;

    const idx = this.fonts.length;
    this.fonts.push(xml);
    this.cache.set(key, idx);
    return idx;
  }

  public renderXml(): string {
    return `<fonts count="${this.fonts.length}">${this.fonts.join("")}</fonts>`;
  }
}
