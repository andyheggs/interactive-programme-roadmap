import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeFilters, ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import { clamp, formatDate, parseDate } from "./dateUtils";

type PdfExportOptions = {
  schedule: ProgrammeSchedule;
  items: ProgrammeItem[];
  viewLabel: string;
  filters: ProgrammeFilters;
  dateWindowLabel: string;
  baselineNumber: number;
};

type TableRow = Array<string | number>;
type Rgb = [number, number, number];
type AutoTable = typeof import("jspdf-autotable").default;

const colours: Record<"ink" | "muted" | "green" | "blue" | "amber" | "red" | "deep" | "pale" | "line", Rgb> = {
  ink: [28, 38, 33],
  muted: [91, 105, 96],
  green: [46, 125, 85],
  blue: [61, 120, 169],
  amber: [197, 139, 40],
  red: [179, 58, 50],
  deep: [33, 76, 67],
  pale: [239, 244, 241],
  line: [199, 209, 203],
};

const streamPalette: Rgb[] = [
  [204, 141, 36],
  [49, 94, 156],
  [36, 126, 84],
  [147, 80, 154],
  [196, 75, 63],
  [52, 136, 145],
  [121, 104, 42],
  [82, 92, 122],
  [172, 92, 42],
  [78, 132, 62],
];

function setFill(doc: JsPDF, colour: Rgb) {
  doc.setFillColor(...colour);
}

function safe(value?: string | number): string {
  if (value === undefined || value === null || value === "") return "-";
  return String(value);
}

function fileSlug(value: string): string {
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "programme-roadmap";
}

function streamColour(stream?: string): Rgb {
  const key = stream || "Unassigned";
  const index = [...key].reduce((total, char) => total + char.charCodeAt(0), 0) % streamPalette.length;
  return streamPalette[index];
}

function addHeader(doc: JsPDF, schedule: ProgrammeSchedule, viewLabel: string) {
  setFill(doc, colours.deep);
  doc.rect(0, 0, 297, 25, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text(schedule.title || "Programme roadmap", 12, 11);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text(`${viewLabel} report`, 12, 18);
  doc.text(`Generated ${formatDate(new Date().toISOString())}`, 250, 18, { align: "right" });
  doc.setTextColor(...colours.ink);
}

function addSectionTitle(doc: JsPDF, title: string, y: number) {
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(...colours.ink);
  doc.text(title, 12, y);
  doc.setDrawColor(...colours.line);
  doc.line(12, y + 3, 285, y + 3);
}

function table(doc: JsPDF, autoTable: AutoTable, y: number, head: string[], body: TableRow[]) {
  autoTable(doc, {
    startY: y,
    head: [head],
    body,
    theme: "grid",
    styles: {
      font: "helvetica",
      fontSize: 7.5,
      cellPadding: 2,
      overflow: "linebreak",
      textColor: colours.ink,
      lineColor: colours.line,
      lineWidth: 0.1,
    },
    headStyles: {
      fillColor: colours.deep,
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    alternateRowStyles: {
      fillColor: [248, 250, 248],
    },
    margin: { left: 12, right: 12 },
  });
}

function summaryRows(schedule: ProgrammeSchedule, items: ProgrammeItem[]): TableRow[] {
  return [
    ["Programme start", formatDate(schedule.startDate), "Items in report", items.length],
    ["Programme finish", formatDate(schedule.finishDate), "Roadmap milestones", items.filter((item) => item.roadmapMilestone).length],
    ["Status date", formatDate(schedule.statusDate), "Critical open items", items.filter((item) => item.isCritical && item.status !== "complete").length],
    ["Source file", safe(schedule.sourceFileName), "Delayed items", items.filter((item) => item.delayDays && item.delayDays > 0).length],
  ];
}

function filterRows(filters: ProgrammeFilters, viewLabel: string, dateWindowLabel: string, baselineNumber: number): TableRow[] {
  const rows: TableRow[] = [
    ["View", viewLabel, "Date window", dateWindowLabel],
    ["Baseline", `Baseline ${baselineNumber}`, "Search", filters.search || "-"],
  ];
  [
    ["Stream", filters.stream],
    ["Roadmap view", filters.roadmapView],
    ["Milestone type", filters.milestoneType],
    ["Approval body", filters.approvalBody],
    ["Version", filters.version],
    ["Visibility", filters.visibility],
    ["Status", filters.status],
  ].forEach(([label, value]) => {
    if (value !== "all") rows.push([label, value, "", ""]);
  });
  const flags = [
    filters.criticalOnly ? "Critical only" : "",
    filters.roadmapOnly ? "Roadmap milestones only" : "",
    filters.delayedOnly ? "Delayed only" : "",
  ].filter(Boolean);
  if (flags.length) rows.push(["Additional filters", flags.join(", "), "", ""]);
  return rows;
}

function milestoneRows(items: ProgrammeItem[]): TableRow[] {
  return items
    .filter((item) => item.isMilestone || item.roadmapMilestone)
    .sort((a, b) => (parseDate(a.finishDate)?.getTime() ?? 0) - (parseDate(b.finishDate)?.getTime() ?? 0))
    .slice(0, 24)
    .map((item) => [
      formatDate(item.finishDate),
      item.name,
      safe(item.stream),
      safe(item.milestoneType),
      item.roadmapMilestone ? "Yes" : "No",
      item.delayDays && item.delayDays > 0 ? `+${item.delayDays}d` : "-",
    ]);
}

function riskRows(items: ProgrammeItem[]): TableRow[] {
  return items
    .filter((item) => item.isCritical || (item.delayDays && item.delayDays > 0) || item.status === "late" || item.status === "at-risk")
    .sort((a, b) => (b.delayDays ?? 0) - (a.delayDays ?? 0))
    .slice(0, 24)
    .map((item) => [
      item.name,
      safe(item.stream),
      formatDate(item.finishDate),
      item.status,
      item.isCritical ? "Yes" : "No",
      item.delayDays && item.delayDays > 0 ? `+${item.delayDays}d` : "-",
    ]);
}

function governanceRows(items: ProgrammeItem[]): TableRow[] {
  return items
    .filter((item) => item.approvalBody && (item.isMilestone || item.roadmapMilestone || item.milestoneType === "Approval"))
    .sort((a, b) => `${a.approvalBody}`.localeCompare(`${b.approvalBody}`) || (parseDate(a.finishDate)?.getTime() ?? 0) - (parseDate(b.finishDate)?.getTime() ?? 0))
    .slice(0, 24)
    .map((item) => [
      safe(item.approvalBody),
      formatDate(item.finishDate),
      item.name,
      safe(item.stream),
      safe(item.visibility),
      item.delayDays && item.delayDays > 0 ? `+${item.delayDays}d` : "-",
    ]);
}

function timelineCandidates(items: ProgrammeItem[], timelineStart: Date): ProgrammeItem[] {
  return items
    .filter((item) => item.roadmapMilestone && item.finishDate)
    .filter((item) => {
      const finish = parseDate(item.finishDate);
      return finish ? finish >= timelineStart : false;
    })
    .sort((a, b) => `${a.stream || "Unassigned"}`.localeCompare(`${b.stream || "Unassigned"}`) || (parseDate(a.finishDate)?.getTime() ?? 0) - (parseDate(b.finishDate)?.getTime() ?? 0));
}

function timelineItems(items: ProgrammeItem[], timelineStart: Date, limit = 30): ProgrammeItem[] {
  return timelineCandidates(items, timelineStart)
    .slice(0, limit);
}

function drawLegendItem(doc: JsPDF, x: number, y: number, label: string, draw: () => void) {
  draw();
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.3);
  doc.setTextColor(...colours.ink);
  doc.text(label, x + 11, y + 1.2);
}

function drawTimelineLegend(doc: JsPDF, y: number, note: string) {
  doc.setFillColor(...colours.pale);
  doc.roundedRect(12, y, 273, 24, 1.5, 1.5, "F");
  doc.setDrawColor(...colours.line);
  doc.roundedRect(12, y, 273, 24, 1.5, 1.5, "S");

  drawLegendItem(doc, 18, y + 7, "Roadmap milestone", () => {
    setFill(doc, colours.amber);
    doc.triangle(22, y + 3.8, 25, y + 6.8, 22, y + 9.8, "F");
    doc.triangle(22, y + 3.8, 19, y + 6.8, 22, y + 9.8, "F");
  });
  drawLegendItem(doc, 72, y + 7, "Stream colour", () => {
    setFill(doc, streamPalette[1]);
    doc.rect(72, y + 3.9, 7, 5.8, "F");
  });
  drawLegendItem(doc, 122, y + 7, "Critical marker", () => {
    doc.setDrawColor(...colours.red);
    doc.circle(126, y + 6.8, 2, "S");
  });
  drawLegendItem(doc, 170, y + 7, "Baseline finish", () => {
    doc.setDrawColor(130, 142, 134);
    doc.line(174, y + 3.4, 174, y + 10.2);
  });

  drawLegendItem(doc, 18, y + 15, "Late / delayed", () => {
    setFill(doc, colours.red);
    doc.roundedRect(18, y + 12.7, 8, 2.8, 0.5, 0.5, "F");
  });
  drawLegendItem(doc, 58, y + 15, "At risk", () => {
    setFill(doc, colours.amber);
    doc.roundedRect(58, y + 12.7, 8, 2.8, 0.5, 0.5, "F");
  });

  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.1);
  doc.setTextColor(...colours.muted);
  doc.text(note, 18, y + 21.4);
}

function drawTimeline(doc: JsPDF, items: ProgrammeItem[], schedule: ProgrammeSchedule, y: number, limit: number, timelineStart: Date) {
  const rows = timelineItems(items, timelineStart, limit);
  if (!rows.length) {
    doc.setFontSize(9);
    doc.setTextColor(...colours.muted);
    doc.text("No forward-looking roadmap milestones are available for this filtered view.", 12, y + 8);
    return;
  }

  const dates = rows
    .flatMap((item) => [parseDate(item.finishDate), parseDate(item.baselineFinish)])
    .concat([timelineStart, parseDate(schedule.finishDate)])
    .filter((date): date is Date => Boolean(date));
  const min = timelineStart;
  const max = new Date(Math.max(...dates.map((date) => date.getTime())));
  const span = Math.max(1, max.getTime() - min.getTime());
  const labelWidth = 84;
  const x0 = 12 + labelWidth;
  const x1 = 285;
  const trackWidth = x1 - x0;
  const position = (value?: string) => {
    const date = parseDate(value);
    if (!date) return x0;
    return x0 + clamp((date.getTime() - min.getTime()) / span, 0, 1) * trackWidth;
  };

  doc.setFont("helvetica", "normal");
  doc.setFontSize(7);
  doc.setTextColor(...colours.muted);
  doc.setFont("helvetica", "bold");
  doc.text(formatDate(min.toISOString()), x0, y);
  doc.text(formatDate(max.toISOString()), x1, y, { align: "right" });
  doc.setDrawColor(...colours.line);
  doc.line(x0, y + 4, x1, y + 4);

  rows.forEach((item, index) => {
    const rowY = y + 10 + index * 5.4;
    const label = item.name.length > 44 ? `${item.name.slice(0, 43)}...` : item.name;
    const colour = streamColour(item.stream);
    doc.setFontSize(6.8);
    setFill(doc, colour);
    doc.rect(12, rowY - 2.2, 2.8, 3.8, "F");
    doc.setTextColor(...colours.ink);
    doc.text(label, 17, rowY + 0.2);
    doc.setFontSize(5.6);
    doc.setTextColor(...colours.muted);
    doc.text(safe(item.stream), 17, rowY + 2.6);
    doc.setDrawColor(230, 235, 231);
    doc.line(x0, rowY, x1, rowY);

    const finish = position(item.finishDate);
    if (item.baselineFinish) {
      const baseline = position(item.baselineFinish);
      doc.setDrawColor(130, 142, 134);
      doc.line(baseline, rowY - 1.8, baseline, rowY + 2.2);
    }
    setFill(doc, colour);
    doc.triangle(finish, rowY - 2.5, finish + 2.5, rowY, finish, rowY + 2.5, "F");
    doc.triangle(finish, rowY - 2.5, finish - 2.5, rowY, finish, rowY + 2.5, "F");
    if (item.delayDays && item.delayDays > 0) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(5.8);
      doc.setTextColor(...colours.red);
      doc.text(`+${item.delayDays}d`, Math.min(x1 - 9, finish + 4), rowY + 1.7);
    }
    if (item.isCritical) {
      doc.setDrawColor(...colours.red);
      doc.circle(finish + 3, rowY, 1.2, "S");
    }
  });
}

export async function exportProgrammePdf({ schedule, items, viewLabel, filters, dateWindowLabel, baselineNumber }: PdfExportOptions) {
  const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  addHeader(doc, schedule, viewLabel);

  addSectionTitle(doc, "Executive Summary", 36);
  table(doc, autoTable, 42, ["Metric", "Value", "Metric", "Value"], summaryRows(schedule, items));
  addSectionTitle(doc, "Report Scope", 83);
  table(doc, autoTable, 89, ["Filter", "Value", "Filter", "Value"], filterRows(filters, viewLabel, dateWindowLabel, baselineNumber));

  doc.addPage();
  addHeader(doc, schedule, viewLabel);
  addSectionTitle(doc, "Simplified Roadmap Timeline", 36);
  const timelineStart = new Date();
  const timelineLimit = 20;
  const timelineCandidateCount = timelineCandidates(items, timelineStart).length;
  const timelineNote = `Timeline starts from report date (${formatDate(timelineStart.toISOString())}); showing ${Math.min(timelineLimit, timelineCandidateCount)} of ${timelineCandidateCount} forward-looking roadmap milestones.`;
  drawTimelineLegend(doc, 43, timelineNote);
  drawTimeline(doc, items, schedule, 78, timelineLimit, timelineStart);

  doc.addPage();
  addHeader(doc, schedule, viewLabel);
  addSectionTitle(doc, "Key Roadmap Milestones", 36);
  table(doc, autoTable, 42, ["Finish", "Milestone", "Stream", "Type", "Roadmap", "Delay"], milestoneRows(items));

  doc.addPage();
  addHeader(doc, schedule, viewLabel);
  addSectionTitle(doc, "Critical And Delayed Items", 36);
  table(doc, autoTable, 42, ["Item", "Stream", "Finish", "Status", "Critical", "Delay"], riskRows(items));

  doc.addPage();
  addHeader(doc, schedule, viewLabel);
  addSectionTitle(doc, "Governance Decisions", 36);
  table(doc, autoTable, 42, ["Approval body", "Finish", "Item", "Stream", "Visibility", "Delay"], governanceRows(items));

  const pageCount = doc.getNumberOfPages();
  for (let page = 1; page <= pageCount; page += 1) {
    doc.setPage(page);
    doc.setFontSize(7);
    doc.setTextColor(...colours.muted);
    doc.text(`Page ${page} of ${pageCount}`, 285, 202, { align: "right" });
  }

  doc.save(`${fileSlug(schedule.title)}-${fileSlug(viewLabel)}-board-pack.pdf`);
}
