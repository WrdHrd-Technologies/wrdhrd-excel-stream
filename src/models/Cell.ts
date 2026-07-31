import { CellStyle } from "./CellStyle";

export type CellValue = string | number | boolean | Date | null;

export interface CellObject {
  value?: CellValue;
  formula?: string;
  style?: CellStyle | number;
}

export type CellInput = CellValue | CellObject;

export interface ColumnOption {
  width?: number;
  hidden?: boolean;
}

export interface FreezePaneOption {
  rows?: number;
  columns?: number;
}

export interface WorksheetOptions {
  columns?: ColumnOption[];
  freezePanes?: FreezePaneOption;
  autoFilterRange?: string;
}
