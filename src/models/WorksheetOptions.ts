export interface PaneOptions {
  /**
   * The number of horizontal rows to lock/freeze at the top of the worksheet view.
   */
  xSplit?: number;

  /**
   * The number of vertical columns to lock/freeze at the left edge of the worksheet view.
   */
  ySplit?: number;

  /**
   * Top-left cell marker index visible in the active scroll split boundary (e.g. "D5").
   */
  topLeftCell?: string;

  /**
   * Active region quadrant pane identifier.
   */
  activePane?: "bottomLeft" | "bottomRight" | "topLeft" | "topRight";
}

export interface WorksheetOptions {
  /**
   * Flag determining sheet layout presentation alignment vector. True turns on Right-To-Left display metrics.
   */
  rightToLeft?: boolean;

  /**
   * Grid lines visibility state inside the execution spreadsheet context layer.
   */
  showGridLines?: boolean;

  /**
   * Default magnification zoom scale factor applied across active sheet contexts (e.g., 100).
   */
  zoomScale?: number;

  /**
   * Structural parameter configurations defining layout view window freezes.
   */
  freezePane?: PaneOptions;

  /**
   * Encryption security protection flags controlling editing permissions.
   */
  protection?: {
    enabled: boolean;
    password?: string;
    selectLockedCells?: boolean;
    selectUnlockedCells?: boolean;
    formatCells?: boolean;
    formatColumns?: boolean;
    formatRows?: boolean;
    insertColumns?: boolean;
    insertRows?: boolean;
    deleteColumns?: boolean;
    deleteRows?: boolean;
  };
}
