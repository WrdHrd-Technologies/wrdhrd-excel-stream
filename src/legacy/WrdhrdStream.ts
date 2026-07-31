import archiver from "archiver";
import { PassThrough, Writable } from "stream";

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  color?: string;
  bgColor?: string;
  horizontal?: "left" | "center" | "right" | "fill" | "justify";
  vertical?: "top" | "center" | "bottom" | "justify";
  wrapText?: boolean;
  border?: boolean;
}

export interface CellData {
  value: string | number;
  style?: CellStyle;
}

export class WrdhrdExcelStream {
  private archive: archiver.Archiver;
  private sheets: string[] = [];
  private currentSheetStream: PassThrough | null = null;
  private currentMerges: string[] = [];
  private currentRow: number = 1;
  private colWidths: Map<number, number> = new Map();

  private fonts: string[] = [`<font><sz val="11"/><name val="Calibri"/></font>`];
  private fills: string[] = [
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
  ];
  private cellXfs: string[] = [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`];
  private styleMap: Map<string, number> = new Map();
  private borders: string[] = [
    `<border><left/><right/><top/><bottom/><diagonal/></border>`, // Index 0: No border
    `<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>`, // Index 1: Thin border
  ];

  constructor(outputStream: Writable) {
    this.archive = archiver("zip", { zlib: { level: 9 } });
    this.archive.pipe(outputStream);
  }

  public addSheet(sheetName: string): void {
    if (this.currentSheetStream) {
      this.closeCurrentSheet();
    }

    this.sheets.push(sheetName);
    const sheetIndex = this.sheets.length;
    this.currentRow = 1;
    this.currentMerges = [];
    this.currentSheetStream = new PassThrough();

    this.archive.append(this.currentSheetStream, { name: `xl/worksheets/sheet${sheetIndex}.xml` });
    this.currentSheetStream.write(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>`
    );
  }

  private closeCurrentSheet(): void {
    if (!this.currentSheetStream) return;
    this.currentSheetStream.write(`</sheetData>`);
    if (this.currentMerges.length > 0) {
      this.currentSheetStream.write(
        `<mergeCells count="${this.currentMerges.length}">${this.currentMerges.join("")}</mergeCells>`
      );
    }
    this.currentSheetStream.write(`</worksheet>`);
    this.currentSheetStream.end();
    this.currentSheetStream = null;
  }

  private getColStr(colIndex: number): string {
    let str = "";
    let c = colIndex;
    while (c >= 0) {
      str = String.fromCharCode((c % 26) + 65) + str;
      c = Math.floor(c / 26) - 1;
    }
    return str;
  }

  private registerStyle(style: CellStyle): number {
    if (!style || Object.keys(style).length === 0) return 0;
    const styleKey = JSON.stringify(style);
    if (this.styleMap.has(styleKey)) return this.styleMap.get(styleKey)!;

    let fontXml = `<font>`;
    if (style.bold) fontXml += `<b/>`;
    if (style.italic) fontXml += `<i/>`;
    fontXml += `<sz val="${style.fontSize || 11}"/>`;
    if (style.color) fontXml += `<color rgb="FF${style.color.replace("#", "")}"/>`;
    fontXml += `<name val="Calibri"/></font>`;
    const fontId = this.fonts.length;
    this.fonts.push(fontXml);

    let fillId = 0;
    if (style.bgColor) {
      fillId = this.fills.length;
      this.fills.push(
        `<fill><patternFill patternType="solid"><fgColor rgb="FF${style.bgColor.replace("#", "")}"/></patternFill></fill>`
      );
    }

    const borderId = style.border ? 1 : 0;

    let alignmentXml = "";
    let applyAlignment = 0;
    if (style.horizontal || style.vertical || style.wrapText) {
      const hAlign = style.horizontal ? `horizontal="${style.horizontal}"` : "";
      const vAlign = style.vertical ? `vertical="${style.vertical}"` : `vertical="center"`;
      const wrap = style.wrapText ? `wrapText="1"` : "";

      alignmentXml = `<alignment ${hAlign} ${vAlign} ${wrap}/>`;
      applyAlignment = 1;
    }

    const xfId = this.cellXfs.length;
    this.cellXfs.push(
      `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyFont="1" applyFill="${fillId > 0 ? 1 : 0}" applyBorder="${borderId > 0 ? 1 : 0}" applyAlignment="${applyAlignment}">${alignmentXml}</xf>`
    );

    this.styleMap.set(styleKey, xfId);
    return xfId;
  }

  private escapeXml(unsafe: string): string {
    return unsafe.replace(/[<>&'"]/g, (c) => {
      switch (c) {
        case "<":
          return "&lt;";
        case ">":
          return "&gt;";
        case "&":
          return "&amp;";
        case "'":
          return "&apos;";
        case '"':
          return "&quot;";
        default:
          return c;
      }
    });
  }

  public async writeRow(cells: Array<CellData | null | undefined>): Promise<void> {
    if (!this.currentSheetStream) throw new Error("You must call addSheet() before writing rows.");

    let rowXml = `<row r="${this.currentRow}">`;
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c];
      if (!cell) continue;

      const valStr = String(cell.value);
      const currentMax = this.colWidths.get(c) || 0;
      if (valStr.length > currentMax) {
        this.colWidths.set(c, valStr.length);
      }

      const ref = `${this.getColStr(c)}${this.currentRow}`;
      const styleId = this.registerStyle(cell.style || {});
      const tAttr = typeof cell.value === "string" ? `t="inlineStr"` : ``;

      rowXml += `<c r="${ref}" ${tAttr} s="${styleId}">`;
      if (typeof cell.value === "string") {
        rowXml += `<is><t>${this.escapeXml(cell.value)}</t></is>`;
      } else {
        rowXml += `<v>${cell.value}</v>`;
      }
      rowXml += `</c>`;
    }
    rowXml += `</row>`;
    this.currentRow++;

    const canContinue = this.currentSheetStream.write(rowXml);
    if (!canContinue) {
      return new Promise((resolve) => {
        this.currentSheetStream!.once("drain", resolve);
      });
    }
    return Promise.resolve();
  }

  public merge(range: string): void {
    if (!this.currentSheetStream) throw new Error("You must call addSheet() before merging.");
    this.currentMerges.push(`<mergeCell ref="${range}"/>`);
  }

  public mergeRange(startCol: number, startRow: number, endCol: number, endRow: number): void {
    const sRow = startRow + 1;
    const eRow = endRow + 1;

    const range = `${this.getColStr(startCol)}${sRow}:${this.getColStr(endCol)}${eRow}`;
    this.currentMerges.push(`<mergeCell ref="${range}"/>`);
  }

  private writeGlobalMetadata(): void {
    let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;
    for (let i = 1; i <= this.sheets.length; i++) {
      contentTypes += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    }
    contentTypes += `</Types>`;
    this.archive.append(contentTypes, { name: "[Content_Types].xml" });

    let workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`;
    for (let i = 0; i < this.sheets.length; i++) {
      workbook += `<sheet name="${this.escapeXml(this.sheets[i])}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
    }
    workbook += `</sheets></workbook>`;
    this.archive.append(workbook, { name: "xl/workbook.xml" });

    let workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (let i = 0; i < this.sheets.length; i++) {
      workbookRels += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
    }
    workbookRels += `<Relationship Id="rId${this.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    this.archive.append(workbookRels, { name: "xl/_rels/workbook.xml.rels" });

    this.archive.append(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      { name: "_rels/.rels" }
    );

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                              <fonts count="${this.fonts.length}">${this.fonts.join("")}</fonts>
                              <fills count="${this.fills.length}">${this.fills.join("")}</fills>
                              <borders count="${this.borders.length}">${this.borders.join("")}</borders>
                              <cellXfs count="${this.cellXfs.length}">${this.cellXfs.join("")}</cellXfs>
                          </styleSheet>`;
    this.archive.append(stylesXml, { name: "xl/styles.xml" });
  }

  public async commit(): Promise<void> {
    if (this.sheets.length === 0) throw new Error("No sheets were added.");
    if (this.currentSheetStream) this.closeCurrentSheet();
    this.writeGlobalMetadata();
    return this.archive.finalize();
  }
}
