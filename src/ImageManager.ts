interface EmbeddedImage {
  id: number;
  extension: "png" | "jpeg";
  buffer: Buffer;
  anchor: {
    fromCol: number;
    fromRow: number;
    toCol: number;
    toRow: number;
  };
}

export class ImageManager {
  private images: EmbeddedImage[] = [];
  private counter = 1;

  public addImage(
    buffer: Buffer,
    extension: "png" | "jpeg",
    from: { col: number; row: number },
    to: { col: number; row: number }
  ): number {
    const id = this.counter++;
    this.images.push({
      id,
      extension,
      buffer,
      anchor: { fromCol: from.col, fromRow: from.row, toCol: to.col, toRow: to.row },
    });
    return id;
  }

  public renderDrawingXml(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml +=
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';

    this.images.forEach((img) => {
      xml += "<xdr:twoCellAnchor>";
      xml += "<xdr:from>";
      xml += `<xdr:col>${img.anchor.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff>`;
      xml += `<xdr:row>${img.anchor.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
      xml += "</xdr:from>";
      xml += "<xdr:to>";
      xml += `<xdr:col>${img.anchor.toCol}</xdr:col><xdr:colOff>0</xdr:colOff>`;
      xml += `<xdr:row>${img.anchor.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
      xml += "</xdr:to>";
      xml += "<xdr:pic>";
      xml += `<xdr:nvPicPr><xdr:cNvPr id="${img.id}" name="Picture ${img.id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`;
      xml += `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${img.id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`;
      xml +=
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>';
      xml += "</xdr:pic>";
      xml += "<xdr:clientData/>";
      xml += "</xdr:twoCellAnchor>";
    });

    xml += "</xdr:wsDr>";
    return xml;
  }

  public get allImages(): EmbeddedImage[] {
    return this.images;
  }
}
