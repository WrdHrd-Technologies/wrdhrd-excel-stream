import { PassThrough } from "stream";
import { RowWriter } from "./RowWriter";
import { StyleManager } from "./StyleManager";
import { MergeManager } from "./MergeManager";
import { CellInput, WorksheetOptions } from "./models/Cell";
import { XmlWriter } from "./XmlWriter";

export class Worksheet {
  private entryStream: PassThrough;
  private rowWriter!: RowWriter;
  private xmlWriter: XmlWriter;
  private mergeManager = new MergeManager();
  private hasStartedRows = false;
  private hasEndedRows = false;

  constructor(
    public readonly name: string,
    public readonly id: number,
    private entryStreamFactory: () => PassThrough,
    private styles: StyleManager,
    private options: WorksheetOptions = {}
  ) {
    this.entryStream = this.entryStreamFactory();
    this.xmlWriter = new XmlWriter(this.entryStream);
    this.writeHeader();
  }

  /**
   * Converts a zero-based column index into its corresponding standard Excel alpha reference string (e.g., 0 -> A, 25 -> Z).
   */
  public getColStr(colIndex: number): string {
    let str = "";
    let c = colIndex;
    while (c >= 0) {
      str = String.fromCharCode((c % 26) + 65) + str;
      c = Math.floor(c / 26) - 1;
    }
    return str;
  }

  private writeHeader(): void {
    this.xmlWriter.raw('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    this.xmlWriter
      .startOpen("worksheet")
      .attribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      .attribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
      .closeTag();

    this.xmlWriter.startOpen("dimension").attribute("ref", "A1:D50005").selfClose();

    // FIXED: Strict schema-compliant view block orchestration
    if (this.options.freezePanes) {
      const xSplit = this.options.freezePanes.columns || 0;
      const ySplit = this.options.freezePanes.rows || 0;

      // Dynamically calculate the precise scrolling top-left coordinate boundary cell
      const targetColStr = this.getColStr(xSplit);
      const targetRowStr = String(ySplit + 1);
      const topLeftCell = `${targetColStr}${targetRowStr}`;

      // Determine active target pane allocation windows cleanly
      let activePane: "bottomLeft" | "topRight" | "bottomRight" = "bottomLeft";
      if (xSplit > 0 && ySplit > 0) activePane = "bottomRight";
      else if (xSplit > 0) activePane = "topRight";

      this.xmlWriter.start("sheetViews");
      this.xmlWriter
        .startOpen("sheetView")
        .attribute("tabSelected", this.id === 1 ? "1" : "0")
        .attribute("workbookViewId", "0")
        .closeTag();

      // 1. Output the quadrant split pane structure
      this.xmlWriter
        .startOpen("pane")
        .attribute("xSplit", xSplit)
        .attribute("ySplit", ySplit)
        .attribute("topLeftCell", topLeftCell)
        .attribute("activePane", activePane)
        .attribute("state", "frozen")
        .selfClose();

      // 2. FIXED: Strict Schema Requirement - Append the active workspace pane window selection references
      if (xSplit > 0 && ySplit > 0) {
        this.xmlWriter.startOpen("selection").attribute("pane", "topRight").selfClose();
        this.xmlWriter.startOpen("selection").attribute("pane", "bottomLeft").selfClose();
        this.xmlWriter
          .startOpen("selection")
          .attribute("pane", "bottomRight")
          .attribute("activeCell", topLeftCell)
          .attribute("sqref", topLeftCell)
          .selfClose();
      } else if (xSplit > 0) {
        this.xmlWriter
          .startOpen("selection")
          .attribute("pane", "topRight")
          .attribute("activeCell", topLeftCell)
          .attribute("sqref", topLeftCell)
          .selfClose();
      } else {
        this.xmlWriter
          .startOpen("selection")
          .attribute("pane", "bottomLeft")
          .attribute("activeCell", topLeftCell)
          .attribute("sqref", topLeftCell)
          .selfClose();
      }

      this.xmlWriter.end("sheetView").end("sheetViews");
    }

    if (this.options.columns && this.options.columns.length > 0) {
      this.xmlWriter.start("cols");
      this.options.columns.forEach((col, idx) => {
        const colNum = idx + 1;
        this.xmlWriter
          .startOpen("col")
          .attribute("min", colNum)
          .attribute("max", colNum)
          .attribute("width", col.width || 12)
          .attribute("customWidth", "1")
          .selfClose();
      });
      this.xmlWriter.end("cols");
    }
  }

  public addRow(row: CellInput[]): boolean {
    if (this.hasEndedRows) throw new Error(`Worksheet ${this.name} is already committed.`);
    if (!this.hasStartedRows) {
      this.xmlWriter.start("sheetData");
      this.rowWriter = new RowWriter(this.entryStream, this.styles);
      this.hasStartedRows = true;
    }
    return this.rowWriter.writeRow(row);
  }

  public mergeCells(range: string): void {
    if (this.hasEndedRows) throw new Error("Cannot add merges after committing worksheet.");
    this.mergeManager.addMerge(range);
  }

  public async end(): Promise<void> {
    if (this.hasEndedRows) return;

    if (!this.hasStartedRows) {
      this.xmlWriter.start("sheetData");
      this.rowWriter = new RowWriter(this.entryStream, this.styles);
      this.hasStartedRows = true;
    }

    this.rowWriter.close(this.options?.autoFilterRange);
    await this.mergeManager.writeXml(this.xmlWriter);

    this.xmlWriter.end("worksheet");
    this.hasEndedRows = true;

    await new Promise<void>((resolve) => {
      if (this.entryStream.writableNeedDrain) {
        this.entryStream.once("drain", () => {
          this.entryStream.end(() => resolve());
        });
      } else {
        this.entryStream.end(() => resolve());
      }
    });
  }
}
