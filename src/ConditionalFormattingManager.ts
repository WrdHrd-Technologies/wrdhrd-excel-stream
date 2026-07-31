import { validateA1Range } from "./utils/validation";

export interface ConditionalRule {
  type: "cellIs" | "expression" | "duplicateValues";
  operator?: "lessThan" | "greaterThan" | "equal";
  formula?: string;
  styleId: number;
}

export class ConditionalFormattingManager {
  private cache = new Map<string, ConditionalRule[]>();

  public addRule(range: string, rule: ConditionalRule): void {
    validateA1Range(range);
    if (!this.cache.has(range)) {
      this.cache.set(range, []);
    }
    this.cache.get(range)!.push(rule);
  }

  public renderXml(): string {
    if (this.cache.size === 0) return "";
    let xml = "";

    for (const [range, rules] of this.cache.entries()) {
      xml += `<conditionalFormatting sqref="${range}">`;
      rules.forEach((rule, index) => {
        const opAttr = rule.operator ? ` operator="${rule.operator}"` : "";
        xml += `<cfRule type="${rule.type}" dxfId="${rule.styleId}" priority="${index + 1}"${opAttr}>`;
        if (rule.formula) xml += `<formula>${rule.formula}</formula>`;
        xml += "</cfRule>";
      });
      xml += "</conditionalFormatting>";
    }

    return xml;
  }
}
