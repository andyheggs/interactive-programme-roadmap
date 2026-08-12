export function parseDate(value?: string): Date | undefined {
  if (!value || value.startsWith("0000")) return undefined;
  const ukDate = value.trim().match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:\s|$)/);
  if (ukDate) {
    const year = Number(ukDate[3].length === 2 ? `20${ukDate[3]}` : ukDate[3]);
    const month = Number(ukDate[2]);
    const day = Number(ukDate[1]);
    const parsed = new Date(Date.UTC(year, month - 1, day));
    if (parsed.getUTCFullYear() === year && parsed.getUTCMonth() === month - 1 && parsed.getUTCDate() === day) {
      return parsed;
    }
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? undefined : date;
}

export function formatDate(value?: string): string {
  const date = parseDate(value);
  if (!date) return "Not set";
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(date);
}

export function daysBetween(start?: string, finish?: string): number | undefined {
  const startDate = parseDate(start);
  const finishDate = parseDate(finish);
  if (!startDate || !finishDate) return undefined;
  return Math.round((finishDate.getTime() - startDate.getTime()) / 86_400_000);
}

export function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

export function uniqueSorted(values: Array<string | undefined>): string[] {
  return [...new Set(values.filter((value): value is string => Boolean(value && value.trim())))]
    .sort((a, b) => a.localeCompare(b));
}

export function durationLabel(value?: string): string {
  if (!value) return "Not set";
  const match = value.match(/PT(?:(\d+)H)?(?:(\d+)M)?/);
  if (!match) return value;
  const hours = Number(match[1] ?? 0);
  const minutes = Number(match[2] ?? 0);
  if (hours >= 8) return `${Math.round(hours / 8)}d`;
  if (hours) return `${hours}h`;
  return `${minutes}m`;
}
