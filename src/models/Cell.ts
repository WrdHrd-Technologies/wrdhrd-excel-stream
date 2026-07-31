import { CellStyle } from "./CellStyle";

export type CellValue = string | number | boolean | Date | null;

export interface RichTextRun {
  text: string;
  font?: any;
}

export interface CellData {
  value?: CellValue | RichTextRun[];
  formula?: string;
  hyperlink?: string;
  style?: CellStyle;
}
