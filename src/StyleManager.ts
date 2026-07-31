import { XmlWriter } from "./XmlWriter";
import { CellStyle, Font, Fill, Border, Alignment } from "./models/CellStyle";

export class StyleManager {
  private fonts: string[] = [];
  private fills: string[] = [];
  private borders: string[] = [];
  private numFormats: Map<string, number> = new Map();
  private cellXfs: string[] = [];
  private customNumFmtIdCounter = 164;

  constructor() {
    this.initializeDefaults();
  }

  private initializeDefaults(): void {
    this.registerFont({ size: 11, name: "Calibri", family: 2, scheme: "minor" });
    this.registerFill({ type: "pattern", patternType: "none" });
    this.registerFill({ type: "pattern", patternType: "gray125" });
    this.registerBorder({});
    this.registerXf({ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 });
  }

  public registerFont(font: Font): number {
    const key = JSON.stringify(font);
    let idx = this.fonts.indexOf(key);
    if (idx === -1) {
      idx = this.fonts.length;
      this.fonts.push(key);
    }
    return idx;
  }

  public registerFill(fill: Fill): number {
    const key = JSON.stringify(fill);
    let idx = this.fills.indexOf(key);
    if (idx === -1) {
      idx = this.fills.length;
      this.fills.push(key);
    }
    return idx;
  }

  public registerBorder(border: Border): number {
    const key = JSON.stringify(border);
    let idx = this.borders.indexOf(key);
    if (idx === -1) {
      idx = this.borders.length;
      this.borders.push(key);
    }
    return idx;
  }

  public registerNumberFormat(formatStr: string): number {
    const builtIns: Record<string, number> = {
      General: 0,
      "0": 1,
      "0.00": 2,
      "#,##0": 3,
      "#,##0.00": 4,
    };
    if (builtIns[formatStr] !== undefined) return builtIns[formatStr];

    if (this.numFormats.has(formatStr)) {
      return this.numFormats.get(formatStr)!;
    }
    const newId = this.customNumFmtIdCounter++;
    this.numFormats.set(formatStr, newId);
    return newId;
  }

  private normalizeColor(color: string | undefined): string | undefined {
    if (!color) return undefined;
    const cleanHex = color.replace("#", "");
    return cleanHex.length === 6 ? `FF${cleanHex}` : cleanHex;
  }

  public registerStyle(style: CellStyle): number {
    let finalFont: Font = { size: 11, name: "Calibri" };
    if (style.font) {
      finalFont = { ...style.font, color: this.normalizeColor(style.font.color) };
    }

    let finalFill: Fill | undefined = undefined;
    if (style.fill) {
      finalFill = {
        ...style.fill,
        foregroundColor: this.normalizeColor(style.fill.foregroundColor),
        backgroundColor: this.normalizeColor(style.fill.backgroundColor),
      };
    }

    let finalBorder: Border | undefined = undefined;
    if (style.border && typeof style.border === "object") {
      finalBorder = style.border;
    } else if (style.border === true) {
      finalBorder = {
        left: { style: "thin", color: "FF000000" },
        right: { style: "thin", color: "FF000000" },
        top: { style: "thin", color: "FF000000" },
        bottom: { style: "thin", color: "FF000000" },
      };
    }

    const fontId = this.registerFont(finalFont);
    const fillId = finalFill ? this.registerFill(finalFill) : 0;
    const borderId = finalBorder ? this.registerBorder(finalBorder) : 0;
    const numFmtId = style.numberFormat ? this.registerNumberFormat(style.numberFormat) : 0;

    const key = JSON.stringify({
      fontId,
      fillId,
      borderId,
      numFmtId,
      alignment: style.alignment,
      applyNumberFormat: numFmtId > 0 ? 1 : 0,
      applyFont: fontId > 0 ? 1 : 0,
      applyFill: fillId > 0 ? 1 : 0,
      applyBorder: borderId > 0 ? 1 : 0,
      applyAlignment: style.alignment ? 1 : 0,
    });

    let idx = this.cellXfs.indexOf(key);
    if (idx === -1) {
      idx = this.cellXfs.length;
      this.cellXfs.push(key);
    }
    return idx;
  }

  public writeXml(writer: XmlWriter): void {
    writer.raw('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    writer
      .startOpen("styleSheet")
      .attribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      .closeTag();

    if (this.numFormats.size > 0) {
      writer.startOpen("numFmts").attribute("count", this.numFormats.size).closeTag();
      this.numFormats.forEach((id, formatCode) => {
        writer
          .startOpen("numFmt")
          .attribute("numFmtId", id)
          .attribute("formatCode", formatCode)
          .selfClose();
      });
      writer.end("numFmts");
    }

    writer.startOpen("fonts").attribute("count", this.fonts.length).closeTag();
    for (const fStr of this.fonts) {
      const f: Font = JSON.parse(fStr);
      writer.start("font");
      if (f.bold) writer.startOpen("b").selfClose();
      if (f.italic) writer.startOpen("i").selfClose();
      if (f.size) writer.startOpen("sz").attribute("val", f.size).selfClose();
      if (f.color) writer.startOpen("color").attribute("rgb", f.color).selfClose();
      if (f.name) writer.startOpen("name").attribute("val", f.name).selfClose();
      if (f.family) writer.startOpen("family").attribute("val", f.family).selfClose();
      if (f.scheme) writer.startOpen("scheme").attribute("val", f.scheme).selfClose();
      writer.end("font");
    }
    writer.end("fonts");

    writer.startOpen("fills").attribute("count", this.fills.length).closeTag();
    for (const fillStr of this.fills) {
      const fill: Fill = JSON.parse(fillStr);
      writer.start("fill");
      writer.startOpen("patternFill").attribute("patternType", fill.patternType || "none");
      if (fill.foregroundColor || fill.backgroundColor) {
        writer.closeTag();
        if (fill.foregroundColor)
          writer.startOpen("fgColor").attribute("rgb", fill.foregroundColor).selfClose();
        if (fill.backgroundColor)
          writer.startOpen("bgColor").attribute("rgb", fill.backgroundColor).selfClose();
        writer.end("patternFill");
      } else {
        writer.selfClose();
      }
      writer.end("fill");
    }
    writer.end("fills");

    writer.startOpen("borders").attribute("count", this.borders.length).closeTag();
    for (const bStr of this.borders) {
      const b: Border = JSON.parse(bStr);
      writer.start("border");
      this.writeBorderSide(writer, "left", b.left);
      this.writeBorderSide(writer, "right", b.right);
      this.writeBorderSide(writer, "top", b.top);
      this.writeBorderSide(writer, "bottom", b.bottom);
      this.writeBorderSide(writer, "diagonal", b.diagonal);
      writer.end("border");
    }
    writer.end("borders");

    writer.startOpen("cellStyleXfs").attribute("count", 1).closeTag();
    writer
      .startOpen("xf")
      .attribute("numFmtId", 0)
      .attribute("fontId", 0)
      .attribute("fillId", 0)
      .attribute("borderId", 0)
      .selfClose();
    writer.end("cellStyleXfs");

    writer.startOpen("cellXfs").attribute("count", this.cellXfs.length).closeTag();
    for (const xfStr of this.cellXfs) {
      const xf = JSON.parse(xfStr);
      writer
        .startOpen("xf")
        .attribute("numFmtId", xf.numFmtId)
        .attribute("fontId", xf.fontId)
        .attribute("fillId", xf.fillId)
        .attribute("borderId", xf.borderId)
        .attribute("xfId", 0)
        .attribute("applyNumberFormat", xf.applyNumberFormat ? "1" : undefined)
        .attribute("applyFont", xf.applyFont ? "1" : undefined)
        .attribute("applyFill", xf.applyFill ? "1" : undefined)
        .attribute("applyBorder", xf.applyBorder ? "1" : undefined)
        .attribute("applyAlignment", xf.applyAlignment ? "1" : undefined);

      if (xf.alignment) {
        writer.closeTag();
        writer
          .startOpen("alignment")
          .attribute("horizontal", xf.alignment.horizontal)
          .attribute("vertical", xf.alignment.vertical)
          .attribute("wrapText", xf.alignment.wrapText ? "1" : undefined)
          .selfClose();
        writer.end("xf");
      } else {
        writer.selfClose(); // Ensures pure self-closed <xf ... /> nodes matching old style
      }
    }
    writer.end("cellXfs");

    writer.startOpen("cellStyles").attribute("count", 1).closeTag();
    writer
      .startOpen("cellStyle")
      .attribute("name", "Normal")
      .attribute("xfId", 0)
      .attribute("builtinId", 0)
      .selfClose();
    writer.end("cellStyles");

    writer.startOpen("dxfs").attribute("count", 0).selfClose();
    writer
      .startOpen("tableStyles")
      .attribute("count", 0)
      .attribute("defaultTableStyle", "TableStyleMedium9")
      .attribute("defaultPivotStyle", "PivotStyleLight16")
      .selfClose();

    writer.end("styleSheet");
  }

  private writeBorderSide(writer: XmlWriter, side: string, config: any): void {
    if (!config || !config.style) {
      writer.startOpen(side).selfClose();
    } else {
      writer.startOpen(side).attribute("style", config.style).closeTag();
      if (config.color)
        writer.startOpen("color").attribute("rgb", this.normalizeColor(config.color)).selfClose();
      writer.end(side);
    }
  }

  private registerXf(xf: any): void {
    this.cellXfs.push(JSON.stringify(xf));
  }
}
