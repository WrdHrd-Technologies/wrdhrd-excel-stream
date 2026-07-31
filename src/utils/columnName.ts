export function getColAlpha(colIndex: number): string {
  let str = "";
  let temp = colIndex;
  while (temp >= 0) {
    str = String.fromCharCode((temp % 26) + 65) + str;
    temp = Math.floor(temp / 26) - 1;
  }
  return str;
}

export function getColIndex(alpha: string): number {
  let idx = 0;
  for (let i = 0; i < alpha.length; i++) {
    idx = idx * 26 + (alpha.charCodeAt(i) - 64);
  }
  return idx - 1;
}
