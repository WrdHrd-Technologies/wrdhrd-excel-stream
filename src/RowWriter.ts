import { Writable } from "stream";
import { XmlWriter } from "./XmlWriter";
import { StyleManager } from "./StyleManager";
import { CellInput, CellObject } from "./models/Cell";
import { getCellReference } from "./utils/columnName";
import { dateToExcelNumber } from "./utils/excelDate";

export class RowWriter {
  private xmlWriter: XmlWriter;
  private currentRowIndex = 0;

  constructor(
    private stream: Writable,
    private styles: StyleManager
  ) {
    this.xmlWriter = new XmlWriter(this.stream);
  }

  public writeRow(cells: CellInput[]): boolean {
    this.currentRowIndex++;

    this.xmlWriter.startOpen("row").attribute("r", this.currentRowIndex).closeTag();

    cells.forEach((cellInput, colZeroBasedIdx) => {
      const colIndex = colZeroBasedIdx + 1;
      const cellRef = getCellReference(colIndex, this.currentRowIndex);
      const normalized = this.normalizeCellInput(cellInput);
      this.writeCell(cellRef, normalized);
    });

    this.xmlWriter.end("row");
    return this.stream.writableNeedDrain;
  }

  private normalizeCellInput(input: CellInput): CellObject {
    if (input !== null && typeof input === "object" && !(input instanceof Date)) {
      return input as CellObject;
    }
    return { value: input as any };
  }

  private writeCell(cellRef: string, cell: CellObject): void {
    const { value, formula, style } = cell;

    const styleIndex = typeof style === "number" ? style : 0;

    this.xmlWriter.startOpen("c").attribute("r", cellRef);
    if (styleIndex > 0) this.xmlWriter.attribute("s", styleIndex);

    if (formula) {
      this.xmlWriter.attribute("t", "str").closeTag();
      this.xmlWriter.start("f").text(formula).end("f");
      if (value !== undefined && value !== null) {
        this.xmlWriter.start("v").text(String(value)).end("v");
      }
      this.xmlWriter.end("c");
      return;
    }

    if (value === null || value === undefined) {
      this.xmlWriter.selfClose();
      return;
    }

    if (typeof value === "string") {
      this.xmlWriter.attribute("t", "inlineStr").closeTag();
      this.xmlWriter.start("is");
      this.xmlWriter.startOpen("t");
      if (value.startsWith(" ") || value.endsWith(" ")) {
        this.xmlWriter.attribute("xml:space", "preserve");
      }
      this.xmlWriter.closeTag().text(value).end("t").end("is");
    } else if (typeof value === "number") {
      this.xmlWriter.closeTag();
      this.xmlWriter.start("v").text(value).end("v");
    } else if (typeof value === "boolean") {
      this.xmlWriter.attribute("t", "b").closeTag();
      this.xmlWriter
        .start("v")
        .text(value ? "1" : "0")
        .end("v");
    } else if (value instanceof Date) {
      this.xmlWriter.closeTag();
      const excelDate = dateToExcelNumber(value);
      this.xmlWriter.start("v").text(excelDate).end("v");
    }

    this.xmlWriter.end("c");
  }

  public close(autoFilterRange?: string): void {
    this.xmlWriter.end("sheetData");

    if (autoFilterRange) {
      this.xmlWriter.startOpen("autoFilter").attribute("ref", autoFilterRange).selfClose();
    }
  }
}
