import { Writable } from "stream";
import { ZipWriter } from "./ZipWriter";
import { StyleManager } from "./StyleManager";
import { RelationshipManager } from "./RelationshipManager";
import { Worksheet } from "./Worksheet";
import { WorksheetOptions } from "./models/Cell";
import { XmlWriter } from "./XmlWriter";

export class Workbook {
  private zipWriter: ZipWriter;
  public readonly styles = new StyleManager();
  private worksheets: Worksheet[] = [];
  private globalRels = new RelationshipManager();
  private workbookRels = new RelationshipManager();
  private isCommitted = false;

  constructor(outputStream: Writable) {
    this.zipWriter = new ZipWriter(outputStream);
    this.initializeGlobalRelationships();
  }

  private initializeGlobalRelationships(): void {
    this.globalRels.registerRelationship(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
      "xl/workbook.xml"
    );
  }

  public addWorksheet(name: string, options?: WorksheetOptions): Worksheet {
    if (this.isCommitted) throw new Error("Cannot add worksheet; workbook has been committed.");
    const sheetId = this.worksheets.length + 1;

    this.workbookRels.registerRelationship(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
      `worksheets/sheet${sheetId}.xml`
    );

    const sheetFactory = () =>
      this.zipWriter.appendEntryStream(`xl/worksheets/sheet${sheetId}.xml`);
    const sheet = new Worksheet(name, sheetId, sheetFactory, this.styles, options);
    this.worksheets.push(sheet);
    return sheet;
  }

  private writePromisifiedEntry(
    zipPath: string,
    writeFn: (writer: XmlWriter) => void
  ): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const stream = this.zipWriter.appendEntryStream(zipPath);
      const writer = new XmlWriter(stream);

      try {
        writeFn(writer);
      } catch (err) {
        stream.end();
        return reject(err);
      }

      stream.on("finish", () => resolve());
      stream.on("error", (err) => reject(err));
      stream.end();
    });
  }

  public async commit(): Promise<void> {
    if (this.isCommitted) return;

    for (const sheet of this.worksheets) {
      await sheet.end();
    }

    this.workbookRels.registerRelationship(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      "styles.xml"
    );

    await this.writePromisifiedEntry("_rels/.rels", (writer) => {
      this.globalRels.writeXml(writer);
    });

    await this.writePromisifiedEntry("xl/_rels/workbook.xml.rels", (writer) => {
      this.workbookRels.writeXml(writer);
    });

    await this.writePromisifiedEntry("xl/workbook.xml", (writer) => {
      writer.raw('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
      writer
        .startOpen("workbook")
        .attribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        .attribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        .closeTag();

      writer.start("sheets");
      this.worksheets.forEach((sheet, index) => {
        writer
          .startOpen("sheet")
          .attribute("name", sheet.name)
          .attribute("sheetId", sheet.id)
          .attribute("r:id", `rId${index + 1}`)
          .selfClose();
      });
      writer.end("sheets").end("workbook");
    });

    await this.writePromisifiedEntry("xl/styles.xml", (writer) => {
      this.styles.writeXml(writer);
    });

    await this.writePromisifiedEntry("[Content_Types].xml", (writer) => {
      writer.raw('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
      writer
        .startOpen("Types")
        .attribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
        .closeTag();

      writer
        .startOpen("Default")
        .attribute("Extension", "rels")
        .attribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")
        .selfClose();
      writer
        .startOpen("Default")
        .attribute("Extension", "xml")
        .attribute("ContentType", "application/xml")
        .selfClose();

      writer
        .startOpen("Override")
        .attribute("PartName", "/xl/workbook.xml")
        .attribute(
          "ContentType",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        )
        .selfClose();
      writer
        .startOpen("Override")
        .attribute("PartName", "/xl/styles.xml")
        .attribute(
          "ContentType",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
        )
        .selfClose();

      this.worksheets.forEach((sheet) => {
        writer
          .startOpen("Override")
          .attribute("PartName", `/xl/worksheets/sheet${sheet.id}.xml`)
          .attribute(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
          )
          .selfClose();
      });

      writer.end("Types");
    });

    await this.zipWriter.finalize();
    this.isCommitted = true;
  }
}
