import { Writable } from "stream";
import { Workbook } from "./Workbook";
import { Worksheet } from "./Worksheet";
import { CellInput } from "./models/Cell";
import { CellStyle } from "./models/CellStyle";

/**
 * -----------------------------------------------------------------
 * 1. LEGACY BACKWARD-COMPATIBILITY LAYER (V1 API Wrapper)
 * -----------------------------------------------------------------
 */

export interface LegacyCellStyle {
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

export interface LegacyCellData {
  value: string | number;
  style?: LegacyCellStyle;
}

export class WrdhrdExcelStream {
  private workbook: Workbook;
  private currentWorksheet: Worksheet | null = null;
  private activeSheetOptions: { columns?: { width: number }[] } = {};

  constructor(outputStream: Writable) {
    // Map directly to your high-performance, O(1) multi-pass streaming core
    this.workbook = new Workbook(outputStream);
  }

  public addSheet(sheetName: string): void {
    // Re-initialize local option states exactly like legacy architecture behavior
    this.activeSheetOptions = { columns: [] };
    this.currentWorksheet = this.workbook.addWorksheet(sheetName, this.activeSheetOptions);
  }

  public async writeRow(cells: Array<LegacyCellData | null | undefined>): Promise<void> {
    if (!this.currentWorksheet) {
      throw new Error("You must call addSheet() before writing rows.");
    }

    const modernCells: CellInput[] = cells.map((cell, colIndex) => {
      if (!cell) return null;

      const valStr = String(cell.value);
      if (!this.activeSheetOptions.columns) this.activeSheetOptions.columns = [];
      if (!this.activeSheetOptions.columns[colIndex]) {
        this.activeSheetOptions.columns[colIndex] = { width: 12 };
      }
      if (valStr.length > this.activeSheetOptions.columns[colIndex].width) {
        this.activeSheetOptions.columns[colIndex].width = valStr.length;
      }

      if (!cell.style) return cell.value;

      const modernStyle: CellStyle = {};

      if (cell.style.bold || cell.style.italic || cell.style.fontSize || cell.style.color) {
        modernStyle.font = {
          bold: cell.style.bold,
          italic: cell.style.italic,
          size: cell.style.fontSize,
          color: cell.style.color ? `FF${cell.style.color.replace("#", "")}` : undefined,
        };
      }

      if (cell.style.bgColor) {
        modernStyle.fill = {
          type: "pattern",
          patternType: "solid",
          foregroundColor: `FF${cell.style.bgColor.replace("#", "")}`,
        };
      }

      if (cell.style.border) {
        modernStyle.border = {
          left: { style: "thin", color: "FF000000" },
          right: { style: "thin", color: "FF000000" },
          top: { style: "thin", color: "FF000000" },
          bottom: { style: "thin", color: "FF000000" },
        };
      }

      if (cell.style.horizontal || cell.style.vertical || cell.style.wrapText) {
        modernStyle.alignment = {
          horizontal: cell.style.horizontal === "fill" ? undefined : cell.style.horizontal,
          vertical: cell.style.vertical,
          wrapText: cell.style.wrapText,
        };
      }

      return {
        value: cell.value,
        style: modernStyle,
      };
    });

    const isBackpressured = this.currentWorksheet.addRow(modernCells);
    if (isBackpressured) {
      return new Promise<void>((resolve) => {
        this.currentWorksheet!["entryStream"].once("drain", () => resolve());
      });
    }
    return Promise.resolve();
  }

  public merge(range: string): void {
    if (!this.currentWorksheet) throw new Error("You must call addSheet() before merging.");
    this.currentWorksheet.mergeCells(range);
  }

  public mergeRange(startCol: number, startRow: number, endCol: number, endRow: number): void {
    if (!this.currentWorksheet) throw new Error("You must call addSheet() before merging.");

    const getColStr = (colIndex: number): string => {
      let str = "";
      let c = colIndex;
      while (c >= 0) {
        str = String.fromCharCode((c % 26) + 65) + str;
        c = Math.floor(c / 26) - 1;
      }
      return str;
    };

    const range = `${getColStr(startCol)}${startRow + 1}:${getColStr(endCol)}${endRow + 1}`;
    this.currentWorksheet.mergeCells(range);
  }

  public async commit(): Promise<void> {
    return this.workbook.commit();
  }
}

/**
 * -----------------------------------------------------------------
 * 2. MODERN V2 ARCHITECTURE EXPORTS (For New Projects)
 * -----------------------------------------------------------------
 */
export { Workbook } from "./Workbook";
export { Worksheet } from "./Worksheet";
export {
  WorksheetOptions,
  ColumnOption,
  FreezePaneOption,
  CellInput,
  CellObject,
  CellValue,
} from "./models/Cell";
export { CellStyle, Font, Fill, Border, BorderDetail, Alignment } from "./models/CellStyle";
export { getColumnName, getCellReference } from "./utils/columnName";
