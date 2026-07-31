export function getColumnName(index: number): string {
  let name = "";
  let idx = index;
  while (idx > 0) {
    const modulo = (idx - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    idx = Math.floor((idx - 1 - modulo) / 26);
  }
  return name;
}

export function getCellReference(colIndex: number, rowIndex: number): string {
  return `${getColumnName(colIndex)}${rowIndex}`;
}
