import archiver from "archiver";
import { Writable, PassThrough } from "stream";

export class ZipWriter {
  private archive: archiver.Archiver;

  constructor(outputStream: Writable) {
    this.archive = archiver("zip", { zlib: { level: 6 } });
    this.archive.pipe(outputStream);
  }

  public appendStream(name: string): PassThrough {
    const pt = new PassThrough();
    this.archive.append(pt, { name });
    return pt;
  }

  public appendBuffer(name: string, buffer: Buffer | string): void {
    this.archive.append(buffer, { name });
  }

  public async finalize(): Promise<void> {
    await this.archive.finalize();
  }

  public destroy(error?: Error): void {
    this.archive.destroy(error);
  }
}
