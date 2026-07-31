import { XmlWriter } from "./XmlWriter";

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: "External";
}

export class RelationshipManager {
  private rels: Relationship[] = [];
  private counter = 1;

  public registerRelationship(type: string, target: string, targetMode?: "External"): string {
    const id = `rId${this.counter++}`;
    this.rels.push({ id, type, target, targetMode });
    return id;
  }

  public writeXml(writer: XmlWriter): void {
    writer.raw('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');

    // 1. Open the structural parent envelope using startOpen() to allow inline attributes
    writer
      .startOpen("Relationships")
      .attribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
      .closeTag(); // Explicitly seal the opening tag bracket with '>'

    for (const rel of this.rels) {
      // 2. Open individual nodes using startOpen() to write attributes cleanly inline
      writer
        .startOpen("Relationship")
        .attribute("Id", rel.id)
        .attribute("Type", rel.type)
        .attribute("Target", rel.target);

      if (rel.targetMode) {
        writer.attribute("TargetMode", rel.targetMode);
      }

      // 3. Since relationships do not contain nested children, render as a literal self-closing tag ' />'
      writer.selfClose();
    }

    // 4. Terminate the complete collection envelope using the stateless literal end tag
    writer.end("Relationships");
  }
}
