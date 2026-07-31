import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { XmlWriter } from "./XmlWriter";

export class MergeManager {
  private count = 0;
  private tempFilePath: string;
  private tempWriteStream: fs.WriteStream;

  constructor() {
    this.tempFilePath = path.join(
      os.tmpdir(),
      `excel_stream_merges_${Date.now()}_${Math.random().toString(36).slice(2)}.xml`
    );
    this.tempWriteStream = fs.createWriteStream(this.tempFilePath);
  }

  public addMerge(range: string): void {
    if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(range)) {
      throw new Error(`Invalid Excel range format: ${range}. Must match 'A1:B2'.`);
    }
    this.count++;
    this.tempWriteStream.write(`<mergeCell ref="${range}" />`);
  }

  public async writeXml(writer: XmlWriter): Promise<void> {
    if (this.count === 0) {
      this.cleanup();
      return;
    }

    await new Promise<void>((resolve) => {
      this.tempWriteStream.end(() => resolve());
    });

    writer.startOpen("mergeCells").attribute("count", this.count).closeTag();

    await new Promise<void>((resolve, reject) => {
      const readStream = fs.createReadStream(this.tempFilePath, { encoding: "utf8" });
      readStream.on("data", (chunk) => writer.raw(chunk as string));
      readStream.on("end", () => resolve());
      readStream.on("error", (err) => reject(err));
    });

    writer.end("mergeCells");
    this.cleanup();
  }

  private cleanup(): void {
    try {
      if (fs.existsSync(this.tempFilePath)) {
        fs.unlinkSync(this.tempFilePath);
      }
    } catch {
      // Quiet fail
    }
  }
}
