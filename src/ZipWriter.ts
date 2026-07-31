import archiver, { Archiver } from "archiver";
import { Writable, PassThrough } from "stream";

export class ZipWriter {
  private archive: Archiver;
  private entryPromises: Promise<void>[] = [];

  constructor(private outputStream: Writable) {
    this.archive = archiver("zip", { zlib: { level: 6 } });
    this.archive.pipe(outputStream);
  }

  /**
   * Appends an entry stream and tracks its consumption lifecycle.
   */
  public appendEntryStream(zipPath: string): PassThrough {
    const entryStream = new PassThrough();

    // Wrap the archiver consumption pipeline for this specific stream in a promise
    const entryPromise = new Promise<void>((resolve, reject) => {
      entryStream.on("end", () => {
        // The PassThrough stream data has fully flushed out to the archiver wrapper
        resolve();
      });
      entryStream.on("error", (err) => {
        reject(err);
      });
    });

    this.entryPromises.push(entryPromise);
    this.archive.append(entryStream, { name: zipPath });

    return entryStream;
  }

  /**
   * Complete Compression Finalization Synchronizer
   * Safely waits for all entry streams to drain fully before sealing the archive signatures.
   */
  public async finalize(): Promise<void> {
    // 1. Force the execution scope to wait until all PassThrough entry pipelines
    // are completely drained and acknowledged by the archiver engine instance.
    await Promise.all(this.entryPromises);

    return new Promise<void>((resolve, reject) => {
      let archiveEnded = false;
      let streamFinished = false;

      const checkCompleteness = () => {
        if (archiveEnded && streamFinished) {
          resolve();
        }
      };

      this.archive.on("end", () => {
        archiveEnded = true;
        checkCompleteness();
      });

      this.outputStream.on("finish", () => {
        streamFinished = true;
        checkCompleteness();
      });

      this.outputStream.on("close", () => {
        streamFinished = true;
        archiveEnded = true;
        resolve();
      });

      this.archive.on("error", (err) => reject(err));
      this.outputStream.on("error", (err) => reject(err));

      // 2. Trigger finalization now that byte offsets are stable and predictable
      this.archive.finalize().catch(reject);
    });
  }
}
