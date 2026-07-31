import { CellStyle } from "./CellStyle";

export interface ColumnOptions {
  /**
   * Width of the column (approximate character count visible).
   */
  width?: number;

  /**
   * Set to true to hide the column completely in the view layer.
   */
  hidden?: boolean;

  /**
   * Base styling configuration default applied to all cells instantiated in this column vector.
   */
  style?: CellStyle;

  /**
   * Outline leveling group hierarchy value (0 to 7) for collapsing multi-column configurations.
   */
  outlineLevel?: number;

  /**
   * Flag indicating if the column width configuration is explicitly custom specified.
   */
  customWidth?: boolean;
}

export class Column {
  public width: number;
  public hidden: boolean;
  public style?: CellStyle;
  public outlineLevel: number;

  constructor(options: ColumnOptions = {}) {
    this.width = options.width ?? 15;
    this.hidden = options.hidden ?? false;
    this.style = options.style;
    this.outlineLevel = options.outlineLevel ?? 0;
  }
}
