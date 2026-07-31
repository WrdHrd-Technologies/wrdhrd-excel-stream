import { Writable } from "stream";
import { escapeXml } from "./utils/escapeXml";

export class XmlWriter {
  constructor(private stream: Writable) {}

  public start(tag: string): this {
    this.stream.write(`<${tag}>`);
    return this;
  }

  public startOpen(tag: string): this {
    this.stream.write(`<${tag}`);
    return this;
  }

  public closeTag(): this {
    this.stream.write(">");
    return this;
  }

  public selfClose(): this {
    this.stream.write("/>");
    return this;
  }

  public attribute(name: string, value: string | number | boolean | null | undefined): this {
    if (value !== null && value !== undefined) {
      this.stream.write(` ${name}="${escapeXml(value)}"`);
    }
    return this;
  }

  public end(tag: string): this {
    this.stream.write(`</${tag}>`);
    return this;
  }

  public text(value: string | number | boolean | null | undefined): this {
    if (value !== null && value !== undefined) {
      this.stream.write(escapeXml(value));
    }
    return this;
  }

  public raw(xmlFragment: string): this {
    this.stream.write(xmlFragment);
    return this;
  }
}
