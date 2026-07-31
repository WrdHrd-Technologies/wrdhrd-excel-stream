import { PassThrough } from "stream";
import { RowWriter } from "./RowWriter";
import { StyleManager } from "./StyleManager";
import { SharedStringManager } from "./SharedStringManager";
import { RelationshipManager } from "./RelationshipManager";
import { HyperlinkManager } from "./HyperlinkManager";
import { MergeManager } from "./MergeManager";
import { CellData, CellValue } from "./models/Cell";

export class Worksheet {
  public rels = new RelationshipManager();
  private hyperlinkManager = new HyperlinkManager();
  private mergeManager = new MergeManager();
  private rowWriter: RowWriter;
  private currentRow = 1;
  private colWidths: number[] = [];

  constructor(
    public readonly name: string,
    private ptStream: PassThrough,
    private styleManager: StyleManager,
    private sstManager: SharedStringManager
  ) {
    this.rowWriter = new RowWriter(
      this.ptStream,
      this.styleManager,
      this.sstManager,
      this.hyperlinkManager,
      this.rels
    );
  }

  public setColumns(widths: number[]): void {
    this.colWidths = widths;
  }

  public beginWorksheetStream(): void {
    let xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    if (this.colWidths.length > 0) {
      xml += "<cols>";
      this.colWidths.forEach((width, index) => {
        xml += `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`;
      });
      xml += "</cols>";
    }

    xml += "<sheetData>";
    this.ptStream.write(xml);
  }

  public addRow(cells: Array<CellData | CellValue>): void {
    this.rowWriter.writeRow(this.currentRow++, cells);
  }

  public mergeCells(range: string): void {
    this.mergeManager.add(range);
  }

  public finalizeWorksheetStream(): void {
    let xml = "</sheetData>";

    if (this.mergeManager.size > 0) xml += this.mergeManager.renderXml();
    if (this.hyperlinkManager.size > 0) xml += this.hyperlinkManager.renderXml();

    if (this.rels.size > 0) {
      xml += `<hyperlinks>`;
    }

    xml += "</worksheet>";
    this.ptStream.write(xml);
    this.ptStream.end();
  }
}
