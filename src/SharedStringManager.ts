import { escapeXml } from "./utils/escapeXml";

export class SharedStringManager {
  private stringMap = new Map<string, number>();
  private stringsArray: string[] = [];
  private totalCount = 0;

  public getIndex(value: string): number {
    this.totalCount++;
    let idx = this.stringMap.get(value);
    if (idx === undefined) {
      idx = this.stringsArray.length;
      this.stringMap.set(value, idx);
      this.stringsArray.push(value);
    }
    return idx;
  }

  public renderXml(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.totalCount}" uniqueCount="${this.stringsArray.length}">`;
    for (let i = 0; i < this.stringsArray.length; i++) {
      xml += `<si><t>${escapeXml(this.stringsArray[i])}</t></si>`;
    }
    xml += "</sst>";
    return xml;
  }
}
