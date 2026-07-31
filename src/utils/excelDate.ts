export function dateToExcelNumber(date: Date): number {
  const epoch = Date.UTC(1899, 11, 30);
  const msPerDay = 86400000;
  const userTimezoneOffset = date.getTimezoneOffset() * 60000;
  const utcDate = date.getTime() - userTimezoneOffset;
  return (utcDate - epoch) / msPerDay;
}
