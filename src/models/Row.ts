import { CellStyle } from "./CellStyle";

export interface RowOptions {
  /**
   * Height of the row measured in points.
   */
  height?: number;

  /**
   * Set to true to collapse/hide the row view slice completely.
   */
  hidden?: boolean;

  /**
   * Custom style layout mapping structure applied across the row slice context.
   */
  style?: CellStyle;

  /**
   * Outline collapsing structural level indicator (0 to 7).
   */
  outlineLevel?: number;

  /**
   * Flag enforcing that row height specifications are explicitly managed.
   */
  customHeight?: boolean;
}

export class Row {
  public height?: number;
  public hidden: boolean;
  public style?: CellStyle;
  public outlineLevel: number;
  public customHeight: boolean;

  constructor(options: RowOptions = {}) {
    this.height = options.height;
    this.hidden = options.hidden ?? false;
    this.style = options.style;
    this.outlineLevel = options.outlineLevel ?? 0;
    this.customHeight = options.customHeight ?? options.height !== undefined;
  }
}
