import { WorksheetOptions } from "./models/WorksheetOptions";

export class WorksheetWriter {
  public static renderHeader(colWidths: number[], options?: WorksheetOptions): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ';
    xml += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    if (options?.rightToLeft || options?.showGridLines === false) {
      xml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"';
      if (options.rightToLeft) xml += ' rightToLeft="1"';
      if (options.showGridLines === false) xml += ' showGridLines="0"';
      xml += "></sheetView></sheetViews>";
    }

    if (colWidths.length > 0) {
      xml += "<cols>";
      colWidths.forEach((width, index) => {
        xml += `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`;
      });
      xml += "</cols>";
    }

    xml += "<sheetData>";
    return xml;
  }
}
