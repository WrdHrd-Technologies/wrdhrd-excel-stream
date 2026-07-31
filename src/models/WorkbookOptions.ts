export interface WorkbookOptions {
  /**
   * Document metadata parameters capturing identity references.
   */
  metadata?: {
    creator?: string;
    lastModifiedBy?: string;
    title?: string;
    subject?: string;
    description?: string;
    company?: string;
    category?: string;
  };

  /**
   * Compression footprint optimization level applied during file assembly (0 to 9).
   */
  compressionLevel?: number;

  /**
   * Enforces writing inline string XML tags rather than using a SharedStrings mapping dictionary.
   * Useful for ultra-low memory optimization tasks where string deduplication processing overhead is skipped.
   */
  useInlineStrings?: boolean;

  /**
   * Active tab identifier visible when the spreadsheet is opened.
   */
  activeTab?: number;
}
