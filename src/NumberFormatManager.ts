export class NumberFormatManager {
  private customFormats = new Map<string, number>();
  private nextId = 164;

  private builtInFormats = new Map<string, number>([
    ["General", 0],
    ["0", 1],
    ["0.00", 2],
    ["#,##0", 3],
    ["#,##0.00", 4],
    ["0%", 9],
    ["0.00%", 10],
    ["yyyy-mm-dd", 14],
    ["hh:mm:ss", 18],
    ["yyyy-mm-dd hh:mm:ss", 22],
  ]);

  public getFormatId(fmt?: string): number {
    if (!fmt) return 0;
    if (this.builtInFormats.has(fmt)) return this.builtInFormats.get(fmt)!;
    if (this.customFormats.has(fmt)) return this.customFormats.get(fmt)!;

    const id = this.nextId++;
    this.customFormats.set(fmt, id);
    return id;
  }

  public renderXml(): string {
    if (this.customFormats.size === 0) return "";
    let xml = `<numFmts count="${this.customFormats.size}">`;
    for (const [fmt, id] of this.customFormats.entries()) {
      xml += `<numFmt numFmtId="${id}" formatCode="${fmt}"/>`;
    }
    xml += "</numFmts>";
    return xml;
  }
}
