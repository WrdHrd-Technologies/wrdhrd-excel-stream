import { Writable } from "stream";
import { escapeXml } from "./utils/escapeXml";

export class XmlWriter {
  private tagStack: string[] = [];

  constructor(private stream: Writable) {}

  public start(tag: string, attributes?: Record<string, string | number | boolean>): this {
    let xml = `<${tag}`;
    if (attributes) {
      for (const [key, value] of Object.entries(attributes)) {
        xml += ` ${key}="${escapeXml(value)}"`;
      }
    }
    xml += ">";
    this.stream.write(xml);
    this.tagStack.push(tag);
    return this;
  }

  public end(): this {
    const tag = this.tagStack.pop();
    if (tag) {
      this.stream.write(`</${tag}>`);
    }
    return this;
  }

  public empty(tag: string, attributes?: Record<string, string | number | boolean>): this {
    let xml = `<${tag}`;
    if (attributes) {
      for (const [key, value] of Object.entries(attributes)) {
        xml += ` ${key}="${escapeXml(value)}"`;
      }
    }
    xml += "/>";
    this.stream.write(xml);
    return this;
  }

  public text(content: string | number | boolean): this {
    this.stream.write(escapeXml(content));
    return this;
  }

  public raw(xml: string): this {
    this.stream.write(xml);
    return this;
  }
}
