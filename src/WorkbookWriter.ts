import { Worksheet } from "./Worksheet";

export class WorkbookWriter {
  public static renderXml(sheets: Worksheet[]): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ';
    xml += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    xml += '<workbookPr defaultThemeVersion="164011"/>';
    xml += "<sheets>";

    sheets.forEach((sheet, index) => {
      // Maps sheet token indicators back down to matching rId index tracks
      xml += `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`;
    });

    xml += "</sheets>";
    xml += '<calcPr calcId="162913"/>';
    xml += "</workbook>";
    return xml;
  }
}
