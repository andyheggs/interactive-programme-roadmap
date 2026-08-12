import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import { clamp, formatDate, parseDate } from "./dateUtils";

export type GanttPdfSection = {
  label: string;
  items: ProgrammeItem[];
};

type DateWindow = {
  start?: Date;
  end?: Date;
  label: string;
};

type Rgb = [number, number, number];

const colours: Record<"ink" | "muted" | "deep" | "line" | "pale" | "green" | "amber" | "red" | "blue" | "summary" | "white", Rgb> = {
  ink: [28, 38, 33],
  muted: [91, 105, 96],
  deep: [33, 76, 67],
  line: [199, 209, 203],
  pale: [243, 247, 245],
  green: [46, 125, 85],
  amber: [255, 138, 0],
  red: [179, 58, 50],
  blue: [61, 120, 169],
  summary: [37, 83, 73],
  white: [255, 255, 255],
};

function fileSlug(value: string): string {
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "programme-gantt";
}

function bySoonest(a?: string, b?: string): number {
  return (parseDate(a)?.getTime() ?? Number.MAX_SAFE_INTEGER) - (parseDate(b)?.getTime() ?? Number.MAX_SAFE_INTEGER);
}

function uniqueItems(items: ProgrammeItem[]): ProgrammeItem[] {
  return [...new Map(items.map((item) => [item.uid, item])).values()];
}

function directPredecessors(item: ProgrammeItem, schedule: ProgrammeSchedule): ProgrammeItem[] {
  const byUid = new Map(schedule.items.map((entry) => [entry.uid, entry]));
  return uniqueItems(item.predecessors
    .map((link) => link.predecessorUid ? byUid.get(link.predecessorUid) : undefined)
    .filter((entry): entry is ProgrammeItem => Boolean(entry?.isActive)));
}

function directSuccessors(item: ProgrammeItem, schedule: ProgrammeSchedule): ProgrammeItem[] {
  const linkedBySuccessor = item.successors
    .map((link) => link.successorUid ? schedule.items.find((entry) => entry.uid === link.successorUid) : undefined)
    .filter((entry): entry is ProgrammeItem => Boolean(entry?.isActive));
  const linkedByPredecessor = schedule.items
    .filter((entry) => entry.isActive && entry.predecessors.some((link) => link.predecessorUid === item.uid));
  return uniqueItems([...linkedBySuccessor, ...linkedByPredecessor]);
}

function ganttBounds(items: ProgrammeItem[], schedule: ProgrammeSchedule, dateWindow: DateWindow) {
  const dates = [
    dateWindow.start ?? parseDate(schedule.startDate),
    dateWindow.end ?? parseDate(schedule.finishDate),
    ...items.flatMap((item) => [parseDate(item.startDate), parseDate(item.finishDate), parseDate(item.baselineFinish)]),
  ].filter((date): date is Date => Boolean(date));
  if (!dates.length) {
    const now = new Date();
    return { min: now, max: new Date(now.getTime() + 90 * 86_400_000), span: 90 * 86_400_000 };
  }
  const min = new Date(Math.min(...dates.map((date) => date.getTime())));
  const max = new Date(Math.max(...dates.map((date) => date.getTime())));
  return { min, max, span: Math.max(1, max.getTime() - min.getTime()) };
}

function position(dateValue: string | undefined, bounds: ReturnType<typeof ganttBounds>): number {
  const date = parseDate(dateValue);
  if (!date) return 0;
  return clamp(((date.getTime() - bounds.min.getTime()) / bounds.span) * 100, 0, 100);
}

function tone(item: ProgrammeItem): keyof typeof colours {
  if (item.status === "late" || item.status === "blocked" || (item.delayDays ?? 0) > 0) return "red";
  if (item.status === "at-risk" || item.decisionRequired || item.externalDependency) return "amber";
  if (item.status === "complete") return "green";
  if (item.isSummary) return "summary";
  return "blue";
}

function setText(doc: JsPDF, colour: Rgb) {
  doc.setTextColor(colour[0], colour[1], colour[2]);
}

function setDraw(doc: JsPDF, colour: Rgb) {
  doc.setDrawColor(colour[0], colour[1], colour[2]);
}

function setFill(doc: JsPDF, colour: Rgb) {
  doc.setFillColor(colour[0], colour[1], colour[2]);
}

function addHeader(doc: JsPDF, schedule: ProgrammeSchedule, dateWindow: DateWindow, label: string) {
  const pageWidth = doc.internal.pageSize.getWidth();
  setFill(doc, colours.deep);
  doc.rect(0, 0, pageWidth, 24, "F");
  setText(doc, colours.white);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.text("Programme Gantt View", 12, 9);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(`${schedule.title} · ${label}`, 12, 17);
  doc.text(`Date window: ${dateWindow.label}`, pageWidth - 12, 9, { align: "right" });
  doc.text(`Generated ${formatDate(new Date().toISOString())}`, pageWidth - 12, 17, { align: "right" });
}

function addFooter(doc: JsPDF) {
  const pageCount = doc.getNumberOfPages();
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  for (let page = 1; page <= pageCount; page += 1) {
    doc.setPage(page);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    setText(doc, colours.muted);
    doc.text(`Page ${page} of ${pageCount}`, pageWidth - 12, pageHeight - 7, { align: "right" });
  }
}

function drawScale(doc: JsPDF, bounds: ReturnType<typeof ganttBounds>, x: number, y: number, w: number) {
  setDraw(doc, colours.line);
  doc.line(x, y + 8, x + w, y + 8);
  const month = new Date(Date.UTC(bounds.min.getUTCFullYear(), bounds.min.getUTCMonth(), 1));
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7);
  setText(doc, colours.muted);
  while (month <= bounds.max) {
    const pct = clamp(((month.getTime() - bounds.min.getTime()) / bounds.span) * 100, 0, 100);
    const tickX = x + (pct / 100) * w;
    doc.line(tickX, y + 5, tickX, y + 11);
    doc.text(new Intl.DateTimeFormat("en-GB", { month: "short", year: "2-digit" }).format(month), tickX + 1, y + 4);
    month.setUTCMonth(month.getUTCMonth() + 1);
  }
}

function drawLegend(doc: JsPDF, x: number, y: number) {
  const items: Array<[keyof typeof colours, string]> = [
    ["summary", "Summary"],
    ["blue", "Planned"],
    ["green", "Complete"],
    ["amber", "At risk"],
    ["red", "Late / blocked"],
    ["blue", "Dependency count shown in row label"],
  ];
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7);
  let cursor = x;
  items.forEach(([key, label]) => {
    setFill(doc, colours[key]);
    doc.roundedRect(cursor, y - 3.2, 5, 3, 0.8, 0.8, "F");
    setText(doc, colours.muted);
    doc.text(label, cursor + 7, y);
    cursor += doc.getTextWidth(label) + 18;
  });
}

function drawSection(doc: JsPDF, schedule: ProgrammeSchedule, section: GanttPdfSection, dateWindow: DateWindow) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12;
  const labelWidth = 78;
  const chartX = margin + labelWidth + 4;
  const chartW = pageWidth - chartX - margin;
  const bounds = ganttBounds(section.items, schedule, dateWindow);
  const rows = section.items
    .slice()
    .sort((a, b) => (a.stream ?? "").localeCompare(b.stream ?? "") || bySoonest(a.startDate ?? a.finishDate, b.startDate ?? b.finishDate))
    .slice(0, 80);
  let y = 34;

  addHeader(doc, schedule, dateWindow, section.label);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  setText(doc, colours.ink);
  doc.text(section.label, margin, y);
  drawLegend(doc, margin + 52, y);
  y += 10;
  drawScale(doc, bounds, chartX, y, chartW);
  y += 15;

  rows.forEach((item) => {
    if (y > pageHeight - 18) {
      doc.addPage();
      addHeader(doc, schedule, dateWindow, section.label);
      y = 34;
      drawScale(doc, bounds, chartX, y, chartW);
      y += 15;
    }
    setDraw(doc, colours.line);
    doc.line(margin, y + 7, pageWidth - margin, y + 7);
    doc.setFont("helvetica", item.isSummary ? "bold" : "normal");
    doc.setFontSize(item.isSummary ? 7.4 : 7);
    setText(doc, item.isSummary ? colours.ink : colours.muted);
    const label = `${item.stream ? `${item.stream}: ` : ""}${item.name}`;
    doc.text(doc.splitTextToSize(label, labelWidth).slice(0, 2), margin, y);
    const predecessorCount = directPredecessors(item, schedule).length;
    const successorCount = directSuccessors(item, schedule).length;
    if (predecessorCount || successorCount) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(5.8);
      setText(doc, colours.blue);
      doc.text(`Pred ${predecessorCount} · Succ ${successorCount}`, margin, y + 5);
    }

    const startPct = item.isMilestone ? position(item.finishDate, bounds) : position(item.startDate ?? item.finishDate, bounds);
    const finishPct = position(item.finishDate ?? item.startDate, bounds);
    const startX = chartX + (Math.min(startPct, finishPct) / 100) * chartW;
    const endX = chartX + (Math.max(startPct, finishPct) / 100) * chartW;
    setFill(doc, colours[tone(item)]);
    if (item.isMilestone) {
      doc.triangle(startX, y - 1, startX + 3, y + 2, startX, y + 5, "F");
      doc.triangle(startX, y - 1, startX - 3, y + 2, startX, y + 5, "F");
    } else {
      doc.roundedRect(startX, y - 1, Math.max(2, endX - startX), item.isSummary ? 5 : 4, 1.2, 1.2, "F");
    }
    if (item.baselineFinish) {
      const baseX = chartX + (position(item.baselineFinish, bounds) / 100) * chartW;
      setDraw(doc, colours.muted);
      doc.line(baseX, y - 3, baseX, y + 6);
    }
    y += 10;
  });

  if (!rows.length) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    setText(doc, colours.muted);
    doc.text("No dated items found for this Gantt level and date window.", margin, y + 8);
  }
}

export async function exportGanttPdf(schedule: ProgrammeSchedule, sections: GanttPdfSection[], dateWindow: DateWindow) {
  const { default: jsPDF } = await import("jspdf");
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  sections.forEach((section, index) => {
    if (index > 0) doc.addPage();
    drawSection(doc, schedule, section, dateWindow);
  });
  addFooter(doc);
  const label = sections.length > 1 ? "all-levels" : fileSlug(sections[0]?.label ?? "gantt");
  doc.save(`${fileSlug(schedule.title)}-gantt-${label}.pdf`);
}
