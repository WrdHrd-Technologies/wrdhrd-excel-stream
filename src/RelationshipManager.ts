interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: "External" | "Internal";
}

export class RelationshipManager {
  private rels: Relationship[] = [];
  private counter = 1;

  public add(type: string, target: string, targetMode?: "External" | "Internal"): string {
    const id = `rId${this.counter++}`;
    this.rels.push({ id, type, target, targetMode });
    return id;
  }

  public renderXml(): string {
    let xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for (const rel of this.rels) {
      const modeAttr = rel.targetMode ? ` TargetMode="${rel.targetMode}"` : "";
      xml += `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"${modeAttr}/>`;
    }
    xml += "</Relationships>";
    return xml;
  }

  public get size(): number {
    return this.rels.length;
  }
}
