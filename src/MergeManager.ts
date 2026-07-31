export class MergeManager {
  private merges: string[] = [];

  public add(range: string): void {
    this.merges.push(`<mergeCell ref="${range}"/>`);
  }

  public renderXml(): string {
    if (this.merges.length === 0) return "";
    return `<mergeCells count="${this.merges.length}">${this.merges.join("")}</mergeCells>`;
  }

  public get size(): number {
    return this.merges.length;
  }
}
