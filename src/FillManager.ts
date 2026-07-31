import { FillStyle } from "./models/CellStyle";

export class FillManager {
  private fills: string[] = [];
  private cache = new Map<string, number>();

  constructor() {
    this.fills.push('<fill><patternFill patternType="none"/></fill>');
    this.fills.push('<fill><patternFill patternType="gray125"/></fill>');
  }

  public getFillIndex(fill?: FillStyle): number {
    if (!fill) return 0;
    const key = JSON.stringify(fill);
    if (this.cache.has(key)) return this.cache.get(key)!;

    let xml = "<fill>";
    if (fill.type === "pattern") {
      xml += `<patternFill patternType="${fill.patternType || "solid"}">`;
      if (fill.fgColor) xml += `<fgColor rgb="${fill.fgColor.replace("#", "")}"/>`;
      if (fill.bgColor) xml += `<bgColor rgb="${fill.bgColor.replace("#", "")}"/>`;
      xml += "</patternFill>";
    }
    xml += "</fill>";

    const idx = this.fills.length;
    this.fills.push(xml);
    this.cache.set(key, idx);
    return idx;
  }

  public renderXml(): string {
    return `<fills count="${this.fills.length}">${this.fills.join("")}</fills>`;
  }
}
