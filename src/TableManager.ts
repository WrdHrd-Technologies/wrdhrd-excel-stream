import { validateA1Range } from "./utils/validation";

export interface TableColumnOptions {
  name: string;
}

export interface TableOptions {
  id: number;
  displayName: string;
  range: string;
  columns: TableColumnOptions[];
  showRowStripes?: boolean;
}

export class TableManager {
  private tables: TableOptions[] = [];

  public addTable(options: TableOptions): void {
    validateA1Range(options.range);
    this.tables.push(options);
  }

  public renderTableXml(table: TableOptions): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" `;
    xml += `id="${table.id}" name="${table.displayName}" displayName="${table.displayName}" `;
    xml += `ref="${table.range}" headerRowCount="1">`;

    xml += `<autoFilter ref="${table.range}"/>`;

    xml += `<tableColumns count="${table.columns.length}">`;
    table.columns.forEach((col, idx) => {
      xml += `<tableColumn id="${idx + 1}" name="${col.name}"/>`;
    });
    xml += "</tableColumns>";

    const stripes = (table.showRowStripes ?? true) ? "1" : "0";
    xml += `<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="${stripes}" showColumnStripes="0"/>`;
    xml += "</table>";
    return xml;
  }

  public get list(): TableOptions[] {
    return this.tables;
  }
}
