import { XmlWriter } from "./XmlWriter";
import { CellData, CellValue } from "./models/Cell";
import { StyleManager } from "./StyleManager";
import { SharedStringManager } from "./SharedStringManager";
import { HyperlinkManager } from "./HyperlinkManager";
import { RelationshipManager } from "./RelationshipManager";
import { getColAlpha } from "./utils/columnName";
import { toExcelDate } from "./utils/excelDate";

export class RowWriter {
  private xmlWriter: XmlWriter;

  constructor(
    private stream: any,
    private styleManager: StyleManager,
    private sstManager: SharedStringManager,
    private hyperlinkManager: HyperlinkManager,
    private rels: RelationshipManager
  ) {
    this.xmlWriter = new XmlWriter(this.stream);
  }

  public writeRow(rowIndex: number, cells: Array<CellData | CellValue>): void {
    this.xmlWriter.raw(`<row r="${rowIndex}">`);

    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      const ref = `${getColAlpha(c)}${rowIndex}`;

      if (cell === null || cell === undefined) {
        continue;
      }

      // Check if it's a raw literal primitive value or explicit config block
      const isObject = typeof cell === "object" && !(cell instanceof Date) && !Array.isArray(cell);
      const val = isObject ? (cell as CellData).value : cell;
      const formula = isObject ? (cell as CellData).formula : undefined;
      const hyperlink = isObject ? (cell as CellData).hyperlink : undefined;
      const style = isObject ? (cell as CellData).style : undefined;

      const styleId = style ? this.styleManager.getStyleStyleId(style) : 0;
      const sAttr = styleId > 0 ? ` s="${styleId}"` : "";

      if (hyperlink) {
        this.hyperlinkManager.add(ref, hyperlink, this.rels);
      }

      if (formula) {
        this.xmlWriter.raw(`<c r="${ref}"${sAttr} t="str"><f>${formula}</f></c>`);
        continue;
      }

      if (val === null || val === undefined) {
        this.xmlWriter.raw(`<c r="${ref}"${sAttr}/>`);
      } else if (typeof val === "number") {
        this.xmlWriter.raw(`<c r="${ref}"${sAttr}><v>${val}</v></c>`);
      } else if (typeof val === "boolean") {
        this.xmlWriter.raw(`<c r="${ref}"${sAttr} t="b"><v>${val ? 1 : 0}</v></c>`);
      } else if (val instanceof Date) {
        const dateStyleId = style
          ? styleId
          : this.styleManager.getStyleStyleId({ numFmt: "yyyy-mm-dd" });
        this.xmlWriter.raw(`<c r="${ref}" s="${dateStyleId}"><v>${toExcelDate(val)}</v></c>`);
      } else {
        // Handle standard standard shared string transformations
        const sstIdx = this.sstManager.getIndex(String(val));
        this.xmlWriter.raw(`<c r="${ref}"${sAttr} t="s"><v>${sstIdx}</v></c>`);
      }
    }

    this.xmlWriter.raw("</row>");
  }
}
