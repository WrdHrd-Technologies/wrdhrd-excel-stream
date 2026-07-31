import { validateA1Range } from "./utils/validation";

export interface DataValidationRule {
  range: string;
  type: "list" | "whole" | "decimal" | "date";
  formula1: string; // Dynamic verification criterion configuration path mapping
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
}

export class DataValidationManager {
  private rules: DataValidationRule[] = [];

  public addRule(rule: DataValidationRule): void {
    validateA1Range(rule.range);
    this.rules.push(rule);
  }

  public renderXml(): string {
    if (this.rules.length === 0) return "";
    let xml = `<dataValidations count="${this.rules.length}">`;

    this.rules.forEach((rule) => {
      const blank = (rule.allowBlank ?? true) ? "1" : "0";
      const showErr = (rule.showErrorMessage ?? true) ? "1" : "0";

      xml += `<dataValidation type="${rule.type}" allowBlank="${blank}" showErrorMessage="${showErr}" sqref="${rule.range}">`;
      xml += `<formula1>${rule.formula1}</formula1>`;
      if (rule.errorTitle) xml += `<errorTitle>${rule.errorTitle}</errorTitle>`;
      if (rule.error) xml += `<error>${rule.error}</error>`;
      xml += "</dataValidation>";
    });

    xml += "</dataValidations>";
    return xml;
  }
}
