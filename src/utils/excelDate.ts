export function toExcelDate(date: Date): number {
  const epoch = Date.UTC(1899, 11, 30);
  const msPerDay = 86400000;
  const targetMs = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    date.getHours(),
    date.getMinutes(),
    date.getSeconds(),
    date.getMilliseconds()
  );
  let diff = (targetMs - epoch) / msPerDay;
  if (diff > 59) diff += 1;
  return diff;
}
