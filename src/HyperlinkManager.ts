import { RelationshipManager } from "./RelationshipManager";

interface HyperlinkRecord {
  cellRef: string;
  rId: string;
}

export class HyperlinkManager {
  private links: HyperlinkRecord[] = [];

  public add(cellRef: string, target: string, relManager: RelationshipManager): void {
    const rId = relManager.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target,
      "External"
    );
    this.links.push({ cellRef, rId });
  }

  public renderXml(): string {
    if (this.links.length === 0) return "";
    let xml = "<hyperlinks>";
    for (const link of this.links) {
      xml += `<hyperlink ref="${link.cellRef}" r:id="${link.rId}"/>`;
    }
    xml += "</hyperlinks>";
    return xml;
  }

  public get size(): number {
    return this.links.length;
  }
}
