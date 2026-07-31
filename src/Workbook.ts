import { Writable } from "stream";
import { ZipWriter } from "./ZipWriter";
import { StyleManager } from "./StyleManager";
import { SharedStringManager } from "./SharedStringManager";
import { Worksheet } from "./Worksheet";
import { RelationshipManager } from "./RelationshipManager";

export class Workbook {
  private zipWriter: ZipWriter;
  private styleManager = new StyleManager();
  private sstManager = new SharedStringManager();
  private sheets: Worksheet[] = [];
  private globalRels = new RelationshipManager();

  constructor(outputStream: Writable) {
    this.zipWriter = new ZipWriter(outputStream);
  }

  public addWorksheet(name: string): Worksheet {
    const pt = this.zipWriter.appendStream(`xl/worksheets/sheet${this.sheets.length + 1}.xml`);
    const sheet = new Worksheet(name, pt, this.styleManager, this.sstManager);
    this.sheets.push(sheet);

    this.globalRels.add(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
      `worksheets/sheet${this.sheets.length}.xml`
    );

    sheet.beginWorksheetStream();
    return sheet;
  }

  public async commit(): Promise<void> {
    if (this.sheets.length === 0) {
      throw new Error("Cannot finalize workbook generation with 0 attached worksheets.");
    }

    for (const sheet of this.sheets) {
      sheet.finalizeWorksheetStream();
      if (sheet.rels.size > 0) {
        this.zipWriter.appendBuffer(
          `xl/worksheets/_rels/sheet${this.sheets.indexOf(sheet) + 1}.xml.rels`,
          sheet.rels.renderXml()
        );
      }
    }

    this.zipWriter.appendBuffer("xl/styles.xml", this.styleManager.renderXml());
    this.zipWriter.appendBuffer("xl/sharedStrings.xml", this.sstManager.renderXml());
    this.zipWriter.appendBuffer("xl/_rels/workbook.xml.rels", this.globalRels.renderXml());

    this.writeWorkbookXml();
    this.writeContentTypesXml();
    this.writeRootRels();

    await this.zipWriter.finalize();
  }

  private writeWorkbookXml(): void {
    let xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';
    this.sheets.forEach((sheet, index) => {
      xml += `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`;
    });
    xml += "</sheets></workbook>";
    this.zipWriter.appendBuffer("xl/workbook.xml", xml);
  }

  private writeContentTypesXml(): void {
    let xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedstrings+xml"/>';
    this.sheets.forEach((_, index) => {
      xml += `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    });
    xml += "</Types>";
    this.zipWriter.appendBuffer("[Content_Types].xml", xml);
  }

  private writeRootRels(): void {
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    this.zipWriter.appendBuffer("_rels/.rels", xml);
  }
}
