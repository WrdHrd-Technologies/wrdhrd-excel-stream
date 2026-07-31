export interface Font {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string; // AARRGGBB hex representation
  family?: number;
  scheme?: string;
}

export interface Fill {
  type: "pattern";
  patternType?: "none" | "solid" | "gray125" | "mediumGray";
  foregroundColor?: string; // AARRGGBB hex
  backgroundColor?: string; // AARRGGBB hex
}

export interface BorderDetail {
  style?: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";
  color?: string; // AARRGGBB hex
}

export interface Border {
  left?: BorderDetail;
  right?: BorderDetail;
  top?: BorderDetail;
  bottom?: BorderDetail;
  diagonal?: BorderDetail;
}

export interface Alignment {
  horizontal?: "left" | "center" | "right" | "justify";
  vertical?: "top" | "center" | "bottom" | "justify";
  wrapText?: boolean;
}

/**
 * Modern Unified CellStyle Interface
 * Enforces production-grade structural properties while preserving
 * seamless backward compatibility for flatter legacy formatting parameters.
 */
export interface CellStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numberFormat?: string;
}
