import { CellStyle } from "./models/CellStyle";
import { FontManager } from "./FontManager";
import { FillManager } from "./FillManager";
import { BorderManager } from "./BorderManager";
import { NumberFormatManager } from "./NumberFormatManager";

export class StyleManager {
  public fonts = new FontManager();
  public fills = new FillManager();
  public borders = new BorderManager();
  public numFmts = new NumberFormatManager();

  private xfs: string[] = [];
  private cache = new Map<string, number>();

  constructor() {
    this.xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
  }

  public getStyleStyleId(style?: CellStyle): number {
    if (!style) return 0;
    const key = JSON.stringify(style);
    if (this.cache.has(key)) return this.cache.get(key)!;

    const fontId = this.fonts.getFontIndex(style.font);
    const fillId = this.fills.getFillIndex(style.fill);
    const borderId = this.borders.getBorderIndex(style.border);
    const numFmtId = this.numFmts.getFormatId(style.numFmt);

    let alignXml = "";
    let hasAlign = false;
    if (style.horizontalAlignment || style.verticalAlignment || style.wrapText) {
      hasAlign = true;
      const h = style.horizontalAlignment ? ` horizontal="${style.horizontalAlignment}"` : "";
      const v = style.verticalAlignment ? ` vertical="${style.verticalAlignment}"` : "";
      const w = style.wrapText ? ` wrapText="1"` : "";
      alignXml = `<alignment${h}${v}${w}/>`;
    }

    let xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyNumberFormat="${numFmtId > 0 ? 1 : 0}" applyFont="${fontId > 0 ? 1 : 0}" applyFill="${fillId > 0 ? 1 : 0}" applyBorder="${borderId > 0 ? 1 : 0}" applyAlignment="${hasAlign ? 1 : 0}">`;
    if (hasAlign) xml += alignXml;
    xml += "</xf>";

    const idx = this.xfs.length;
    this.xfs.push(xml);
    this.cache.set(key, idx);
    return idx;
  }

  public renderXml(): string {
    let xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    xml += this.numFmts.renderXml();
    xml += this.fonts.renderXml();
    xml += this.fills.renderXml();
    xml += this.borders.renderXml();
    xml += `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`;
    xml += `<cellXfs count="${this.xfs.length}">${this.xfs.join("")}</cellXfs>`;
    xml += `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`;
    xml += '<dxfs count="0"/><tableStyles count="0"/>';
    xml += "</styleSheet>";
    return xml;
  }
}
