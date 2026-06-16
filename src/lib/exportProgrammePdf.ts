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

function timelineItems(items: ProgrammeItem[]): ProgrammeItem[] {
  return items
    .filter((item) => item.startDate && item.finishDate && (item.isSummary || item.roadmapMilestone || item.isMilestone || item.isCritical))
    .sort((a, b) => (parseDate(a.startDate)?.getTime() ?? 0) - (parseDate(b.startDate)?.getTime() ?? 0))
    .slice(0, 30);
}

function drawTimeline(doc: JsPDF, items: ProgrammeItem[], schedule: ProgrammeSchedule, y: number) {
  const rows = timelineItems(items);
  if (!rows.length) {
    doc.setFontSize(9);
    doc.setTextColor(...colours.muted);
    doc.text("No dated milestone, summary or critical items are available for this filtered view.", 12, y + 8);
    return;
  }

  const dates = rows
    .flatMap((item) => [parseDate(item.startDate), parseDate(item.finishDate), parseDate(item.baselineFinish)])
    .concat([parseDate(schedule.startDate), parseDate(schedule.finishDate)])
    .filter((date): date is Date => Boolean(date));
  const min = new Date(Math.min(...dates.map((date) => date.getTime())));
  const max = new Date(Math.max(...dates.map((date) => date.getTime())));
  const span = Math.max(1, max.getTime() - min.getTime());
  const labelWidth = 72;
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
  doc.text(formatDate(min.toISOString()), x0, y);
  doc.text(formatDate(max.toISOString()), x1, y, { align: "right" });
  doc.setDrawColor(...colours.line);
  doc.line(x0, y + 4, x1, y + 4);

  rows.forEach((item, index) => {
    const rowY = y + 10 + index * 5.4;
    const label = item.name.length > 40 ? `${item.name.slice(0, 39)}...` : item.name;
    doc.setFontSize(6.8);
    doc.setTextColor(...colours.ink);
    doc.text(label, 12, rowY + 1.2);
    doc.setDrawColor(230, 235, 231);
    doc.line(x0, rowY, x1, rowY);

    const start = position(item.startDate);
    const finish = position(item.finishDate);
    if (item.baselineFinish) {
      const baseline = position(item.baselineFinish);
      doc.setDrawColor(130, 142, 134);
      doc.line(baseline, rowY - 1.8, baseline, rowY + 2.2);
    }
    if (item.isMilestone) {
      setFill(doc, item.roadmapMilestone ? colours.amber : colours.blue);
      doc.triangle(finish, rowY - 2.2, finish + 2.2, rowY, finish, rowY + 2.2, "F");
      doc.triangle(finish, rowY - 2.2, finish - 2.2, rowY, finish, rowY + 2.2, "F");
    } else {
      setFill(doc, item.status === "late" ? colours.red : item.status === "at-risk" ? colours.amber : item.isSummary ? colours.deep : colours.blue);
      doc.roundedRect(Math.min(start, finish), rowY - 1.8, Math.max(2.5, Math.abs(finish - start)), item.isSummary ? 3.4 : 2.4, 0.7, 0.7, "F");
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

  addSectionTitle(doc, "Simplified Roadmap Timeline", 132);
  drawTimeline(doc, items, schedule, 140);

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
