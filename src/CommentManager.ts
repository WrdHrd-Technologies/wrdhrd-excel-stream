interface CellComment {
  cellRef: string;
  author: string;
  text: string;
}

export class CommentManager {
  private comments: CellComment[] = [];

  public addComment(cellRef: string, text: string, author: string = "System"): void {
    this.comments.push({ cellRef, author, text });
  }

  public renderCommentsXml(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    xml += "<authors><author>System</author></authors>";
    xml += "<commentList>";

    this.comments.forEach((c) => {
      xml += `<comment ref="${c.cellRef}" authorId="0">`;
      xml += `<text><r><rPr><sz val="9"/><rFont val="Calibri"/></rPr><t>${c.text}</t></r></text>`;
      xml += "</comment>";
    });

    xml += "</commentList></comments>";
    return xml;
  }

  public get size(): number {
    return this.comments.length;
  }
}
