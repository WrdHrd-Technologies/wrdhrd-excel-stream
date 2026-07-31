export class ContentTypesWriter {
  public static renderXml(worksheetCount: number): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';

    // Default system extension registrations
    xml +=
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
    xml += '<Default Extension="xml" ContentType="application/xml"/>';

    // Core document infrastructure components overrides
    xml +=
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    xml +=
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
    xml +=
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedstrings+xml"/>';

    // Dynamic generation mappings targeting every instantiated worksheet part
    for (let i = 1; i <= worksheetCount; i++) {
      xml += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    }

    xml += "</Types>";
    return xml;
  }
}
