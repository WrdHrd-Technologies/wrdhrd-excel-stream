import { BorderStyle } from "./models/CellStyle";

export class BorderManager {
  private borders: string[] = [];
  private cache = new Map<string, number>();

  constructor() {
    this.borders.push("<border><left/><right/><top/><bottom/><diagonal/></border>");
  }

  public getBorderIndex(border?: BorderStyle): number {
    if (!border) return 0;
    const key = JSON.stringify(border);
    if (this.cache.has(key)) return this.cache.get(key)!;

    let xml = "<border>";
    const sides: Array<"left" | "right" | "top" | "bottom"> = ["left", "right", "top", "bottom"];
    for (const side of sides) {
      const bSide = border[side];
      if (bSide && bSide.style && bSide.style !== "none") {
        xml += `<${side} style="${bSide.style}">`;
        if (bSide.color) xml += `<color rgb="${bSide.color.replace("#", "")}"/>`;
        xml += `</${side}>`;
      } else {
        xml += `<left/>`;
      }
    }
    xml += "<diagonal/></border>";

    const idx = this.borders.length;
    this.borders.push(xml);
    this.cache.set(key, idx);
    return idx;
  }

  public renderXml(): string {
    return `<borders count="${this.borders.length}">${this.borders.join("")}</borders>`;
  }
}
