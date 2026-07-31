export interface FontStyle {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  color?: string;
  name?: string;
}

export interface FillStyle {
  type: "pattern" | "gradient";
  patternType?: "none" | "solid" | "gray125" | "mediumGray";
  fgColor?: string; // ARGB Hex format
  bgColor?: string; // ARGB Hex format
}

export interface BorderSide {
  style?: "none" | "thin" | "medium" | "dashed" | "dotted" | "thick" | "double";
  color?: string; // ARGB Hex format
}

export interface BorderStyle {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
}

export interface CellStyle {
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  numFmt?: string;
  horizontalAlignment?: "left" | "center" | "right" | "fill" | "justify";
  verticalAlignment?: "top" | "center" | "bottom" | "justify";
  wrapText?: boolean;
}
