import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeFilters, ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import { clamp, formatDate, parseDate } from "./dateUtils";

type PdfExportOptions = {
  schedule: ProgrammeSchedule;
  items: ProgrammeItem[];
  viewLabel: string;
  filters: ProgrammeFilters;
  dateWindowLabel: string;
  dateWindowStart?: string;
  dateWindowEnd?: string;
  baselineNumber: number;
  output?: "board-pack" | "poster";
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
  [49, 94, 156],
  [52, 136, 145],
  [36, 126, 84],
  [147, 80, 154],
  [196, 75, 63],
  [121, 104, 42],
  [82, 92, 122],
  [78, 132, 62],
  [90, 118, 150],
  [128, 83, 120],
];

const atRiskColour: Rgb = [138, 86, 28];

const preferredStreamOrder = [
  "legislation",
  "secondary regulations",
  "platform",
  "operations & governance",
  "proof of concept and first adopters",
  "sales and adoption",
  "knowledge",
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
  const lower = key.toLowerCase();
  if (lower.includes("knowledge")) return streamPalette[0];
  if (lower.includes("legislation")) return streamPalette[1];
  if (lower.includes("operation") || lower.includes("governance")) return streamPalette[3];
  if (lower.includes("delivery")) return streamPalette[2];
  if (lower.includes("commercial") || lower.includes("finance")) return streamPalette[4];
  if (lower.includes("communication") || lower.includes("engagement")) return streamPalette[5];
  const index = [...key].reduce((total, char) => total + char.charCodeAt(0), 0) % streamPalette.length;
  return streamPalette[index];
}

function normaliseStream(value: string): string {
  return value.toLowerCase().replace(/\band\b/g, "&").replace(/[^a-z0-9&]+/g, " ").replace(/\s+/g, " ").trim();
}

function streamOrderIndex(stream: string): number {
  const normalised = normaliseStream(stream);
  const index = preferredStreamOrder.findIndex((preferred) => normalised.includes(normaliseStream(preferred)));
  return index === -1 ? preferredStreamOrder.length : index;
}

function reportTitle(schedule: ProgrammeSchedule): string {
  const title = schedule.title || "Programme roadmap";
  return title.toLowerCase().includes("milestone") ? title : `${title} Milestones`;
}

function addHeader(doc: JsPDF, schedule: ProgrammeSchedule, viewLabel: string) {
  const pageWidth = doc.internal.pageSize.getWidth();
  setFill(doc, colours.deep);
  doc.rect(0, 0, pageWidth, 25, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text(reportTitle(schedule), 12, 11);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text(`${viewLabel} report`, 12, 18);
  doc.text(`Generated ${formatDate(new Date().toISOString())}`, pageWidth - 35, 18, { align: "right" });
  doc.setTextColor(...colours.ink);
}

function addSectionTitle(doc: JsPDF, title: string, y: number) {
  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(...colours.ink);
  doc.text(title, 12, y);
  doc.setDrawColor(...colours.line);
  doc.line(12, y + 3, pageWidth - 12, y + 3);
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

function timelineCandidates(items: ProgrammeItem[], timelineStart: Date, timelineEnd?: Date): ProgrammeItem[] {
  return items
    .filter((item) => item.roadmapMilestone && item.finishDate)
    .filter((item) => {
      const finish = parseDate(item.finishDate);
      if (!finish || finish < timelineStart) return false;
      if (timelineEnd && finish > timelineEnd) return false;
      return true;
    })
    .sort((a, b) => streamOrderIndex(a.stream || "Unassigned") - streamOrderIndex(b.stream || "Unassigned") || `${a.stream || "Unassigned"}`.localeCompare(`${b.stream || "Unassigned"}`) || (parseDate(a.finishDate)?.getTime() ?? 0) - (parseDate(b.finishDate)?.getTime() ?? 0));
}

function timelineItems(items: ProgrammeItem[], timelineStart: Date, timelineEnd?: Date, limit = 30): ProgrammeItem[] {
  return timelineCandidates(items, timelineStart, timelineEnd)
    .slice(0, limit);
}

function monthStart(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonths(date: Date, months: number): Date {
  return new Date(date.getFullYear(), date.getMonth() + months, 1);
}

function monthLabel(date: Date): string {
  return date.toLocaleString("en-GB", { month: "short" }).toUpperCase();
}

function tint(colour: Rgb, amount = 0.78): Rgb {
  return colour.map((channel) => Math.round(channel + (255 - channel) * amount)) as Rgb;
}

function drawLegendItem(doc: JsPDF, x: number, y: number, label: string, draw: () => void) {
  draw();
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.3);
  doc.setTextColor(...colours.ink);
  doc.text(label, x + 11, y + 1.2);
}

function drawTimelineLegend(doc: JsPDF, y: number, note: string) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const legendWidth = pageWidth - 24;
  const left = 12;
  doc.setFillColor(...colours.pale);
  doc.roundedRect(left, y, legendWidth, 24, 1.5, 1.5, "F");
  doc.setDrawColor(...colours.line);
  doc.roundedRect(left, y, legendWidth, 24, 1.5, 1.5, "S");

  drawLegendItem(doc, left + 6, y + 7, "Roadmap milestone", () => {
    setFill(doc, colours.amber);
    doc.triangle(left + 10, y + 3.8, left + 13, y + 6.8, left + 10, y + 9.8, "F");
    doc.triangle(left + 10, y + 3.8, left + 7, y + 6.8, left + 10, y + 9.8, "F");
  });

  drawLegendItem(doc, left + 62, y + 7, "Late / delayed", () => {
    setFill(doc, colours.red);
    doc.roundedRect(left + 62, y + 4.7, 8, 2.8, 0.5, 0.5, "F");
  });
  drawLegendItem(doc, left + 108, y + 7, "At risk", () => {
    setFill(doc, atRiskColour);
    doc.roundedRect(left + 108, y + 4.7, 8, 2.8, 0.5, 0.5, "F");
  });
  drawLegendItem(doc, left + 146, y + 7, "Baseline", () => {
    doc.setDrawColor(130, 142, 134);
    doc.setLineDashPattern([1.2, 1.2], 0);
    doc.line(left + 150, y + 3.6, left + 150, y + 10);
    doc.setLineDashPattern([], 0);
  });

  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.1);
  doc.setTextColor(...colours.muted);
  doc.text(note, left + 6, y + 17.4);
}

function drawTimeline(doc: JsPDF, items: ProgrammeItem[], schedule: ProgrammeSchedule, viewLabel: string, title: string, y: number, limit: number, timelineStart: Date, timelineEnd?: Date) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const rows = timelineItems(items, timelineStart, timelineEnd, limit);
  if (!rows.length) {
    doc.setFontSize(9);
    doc.setTextColor(...colours.muted);
    doc.text("No forward-looking roadmap milestones are available for this filtered view.", 12, y + 8);
    return;
  }

  const dates = rows
    .flatMap((item) => [parseDate(item.finishDate), parseDate(item.baselineFinish)])
    .concat([timelineStart, timelineEnd, parseDate(schedule.finishDate)])
    .filter((date): date is Date => Boolean(date));
  const min = monthStart(timelineStart);
  const max = timelineEnd ? addMonths(monthStart(timelineEnd), 1) : addMonths(monthStart(new Date(Math.max(...dates.map((date) => date.getTime())))), 1);
  const span = Math.max(1, max.getTime() - min.getTime());
  const laneLabelWidth = pageWidth > 300 ? 52 : 44;
  const streamKeyWidth = 7;
  const x0 = 12 + streamKeyWidth + laneLabelWidth;
  const x1 = pageWidth - 12;
  const trackWidth = x1 - x0;
  const position = (value?: string) => {
    const date = parseDate(value);
    if (!date) return x0;
    return x0 + clamp((date.getTime() - min.getTime()) / span, 0, 1) * trackWidth;
  };
  const streams = Array.from(
    rows.reduce((groups, item) => {
      const stream = item.stream || "Unassigned";
      const streamItems = [...(groups.get(stream) || []), item]
        .sort((a, b) => (parseDate(a.finishDate)?.getTime() ?? 0) - (parseDate(b.finishDate)?.getTime() ?? 0));
      groups.set(stream, streamItems);
      return groups;
    }, new Map<string, ProgrammeItem[]>())
  ).sort(([aStream, aItems], [bStream, bItems]) => streamOrderIndex(aStream) - streamOrderIndex(bStream) || aStream.localeCompare(bStream) || (parseDate(aItems[0]?.finishDate)?.getTime() ?? 0) - (parseDate(bItems[0]?.finishDate)?.getTime() ?? 0));
  const assignedStreamColours = streams.reduce((map, [stream], index) => {
    map.set(stream, streamPalette[index % streamPalette.length]);
    return map;
  }, new Map<string, Rgb>());
  const headerY = y + 2;
  const monthBandY = y + 9;
  const chartTop = y + 18;
  const chartBottom = pageHeight - 20;
  const chartHeight = Math.max(55, chartBottom - chartTop);
  const laneGap = 3.5;
  const maxLaneHeight = chartHeight;
  const monthCount = Math.max(1, Math.ceil((max.getFullYear() - min.getFullYear()) * 12 + max.getMonth() - min.getMonth()));

  const drawScale = () => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7);
    doc.setTextColor(...colours.muted);
    doc.text(formatDate(timelineStart.toISOString()), x0, headerY);
    doc.text(formatDate(max.toISOString()), x1, headerY, { align: "right" });
    doc.setDrawColor(...colours.line);
    doc.line(x0, headerY + 4, x1, headerY + 4);

    for (let index = 0; index < monthCount; index += 1) {
      const current = addMonths(min, index);
      const next = addMonths(min, index + 1);
      const startX = x0 + ((current.getTime() - min.getTime()) / span) * trackWidth;
      const endX = x0 + ((next.getTime() - min.getTime()) / span) * trackWidth;
      setFill(doc, [27, 99, 125]);
      doc.rect(startX, monthBandY, Math.max(2, endX - startX), 8, "F");
      doc.setDrawColor(9, 49, 62);
      doc.rect(startX, monthBandY, Math.max(2, endX - startX), 8, "S");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(pageWidth > 300 ? 8 : 7);
      doc.setTextColor(255, 255, 255);
      doc.text(monthLabel(current), startX + Math.max(1.5, (endX - startX) / 2), monthBandY + 5.5, { align: "center" });
    }
  };

  const rowHeight = pageWidth > 300 ? 10.2 : 8.2;
  const laneHeightFor = (streamItems: ProgrammeItem[]) => Math.min(maxLaneHeight, Math.max(28, 14 + streamItems.length * rowHeight));
  const startNewTimelinePage = (continued: boolean) => {
    doc.addPage();
    addHeader(doc, schedule, viewLabel);
    addSectionTitle(doc, continued ? `${title} (continued)` : title, 36);
    drawScale();
  };

  drawScale();
  let cursorY = chartTop;
  streams.forEach(([stream, streamItems], streamIndex) => {
    const colour = assignedStreamColours.get(stream) ?? streamColour(stream);
    const laneHeight = laneHeightFor(streamItems);
    if (streamIndex > 0 && cursorY + laneHeight > chartBottom) {
      startNewTimelinePage(true);
      cursorY = chartTop;
    }
    const laneBottom = cursorY + laneHeight;
    const rowStep = Math.max(pageWidth > 300 ? 7.4 : 6, Math.min(rowHeight, (laneHeight - 12) / Math.max(streamItems.length, 1)));

    setFill(doc, tint(colour));
    doc.rect(12, cursorY, pageWidth - 24, laneHeight, "F");
    doc.setDrawColor(8, 44, 53);
    doc.rect(12, cursorY, pageWidth - 24, laneHeight, "S");
    setFill(doc, colour);
    doc.rect(12, cursorY, streamKeyWidth, laneHeight, "F");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(pageWidth > 300 ? 6.8 : 5.7);
    doc.setTextColor(...colours.ink);
    const streamLines = doc.splitTextToSize(stream.toUpperCase(), laneLabelWidth - 7).slice(0, 4);
    const titleY = cursorY + Math.max(7, (laneHeight - streamLines.length * 3.3) / 2);
    doc.text(streamLines, 12 + streamKeyWidth + 3, titleY);

    streamItems.forEach((item, index) => {
      const rowY = cursorY + 7 + index * rowStep;
      if (rowY > laneBottom - 3) return;
      doc.setDrawColor(205, 196, 192);
      doc.line(x0, rowY, x1, rowY);
      const finish = position(item.finishDate);
      if (item.baselineFinish) {
        const baseline = position(item.baselineFinish);
        doc.setDrawColor(130, 142, 134);
        doc.line(baseline, rowY - 2.4, baseline, rowY + 2.6);
      }
      setFill(doc, colours.amber);
      doc.triangle(finish, rowY - 2.7, finish + 2.7, rowY, finish, rowY + 2.7, "F");
      doc.triangle(finish, rowY - 2.7, finish - 2.7, rowY, finish, rowY + 2.7, "F");
      const rawLabel = `${formatDate(item.finishDate)} - ${item.name}`;
      doc.setFont("helvetica", "normal");
      doc.setFontSize(pageWidth > 300 ? 6.8 : 5.5);
      doc.setTextColor(...colours.ink);
      const labelGap = pageWidth > 300 ? 6.8 : 5.4;
      const rightWidth = Math.max(24, x1 - finish - labelGap);
      const leftWidth = Math.max(24, finish - labelGap - x0);
      const preferRight = rightWidth >= Math.min(leftWidth, pageWidth > 300 ? 85 : 48);
      const maxLabelWidth = Math.min(pageWidth > 300 ? 86 : 48, preferRight ? rightWidth : leftWidth);
      const labelLines = doc.splitTextToSize(rawLabel, maxLabelWidth).slice(0, 2);
      if (labelLines.length > 1) {
        const last = labelLines[1];
        if (doc.getTextWidth(last) >= maxLabelWidth - 2 && !last.endsWith("...")) labelLines[1] = `${last.slice(0, Math.max(0, last.length - 4))}...`;
      }
      if (preferRight) {
        doc.text(labelLines, finish + labelGap, rowY - (labelLines.length > 1 ? 1.2 : 0) + 2);
      } else {
        doc.text(labelLines, finish - labelGap, rowY - (labelLines.length > 1 ? 1.2 : 0) + 2, { align: "right" });
      }
      if (item.delayDays && item.delayDays > 0) {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(5.4);
        doc.setTextColor(...colours.red);
        doc.text(`+${item.delayDays}d`, Math.min(x1 - 9, finish + 4), rowY + 4);
      }
    });

    cursorY = laneBottom + laneGap;
  });
}

export async function exportProgrammePdf({ schedule, items, viewLabel, filters, dateWindowLabel, dateWindowStart, dateWindowEnd, baselineNumber, output = "board-pack" }: PdfExportOptions) {
  const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: output === "poster" ? "a3" : "a4" });
  const selectedStart = parseDate(dateWindowStart) ?? parseDate(schedule.startDate) ?? new Date();
  const selectedEnd = parseDate(dateWindowEnd);
  addHeader(doc, schedule, viewLabel);

  if (output === "poster") {
    addSectionTitle(doc, "High Level Roadmap Timeline", 36);
    const timelineStart = selectedStart;
    const timelineLimit = 64;
    const timelineCandidateCount = timelineCandidates(items, timelineStart, selectedEnd).length;
    const timelineNote = `Timeline starts from selected date window (${formatDate(timelineStart.toISOString())}${selectedEnd ? ` to ${formatDate(selectedEnd.toISOString())}` : ""}); showing ${Math.min(timelineLimit, timelineCandidateCount)} of ${timelineCandidateCount} roadmap milestones.`;
    drawTimelineLegend(doc, 43, timelineNote);
    drawTimeline(doc, items, schedule, viewLabel, "High Level Roadmap Timeline", 77, timelineLimit, timelineStart, selectedEnd);
    const pageCount = doc.getNumberOfPages();
    for (let page = 1; page <= pageCount; page += 1) {
      doc.setPage(page);
      doc.setFontSize(7);
      doc.setTextColor(...colours.muted);
      doc.text(`Page ${page} of ${pageCount}`, doc.internal.pageSize.getWidth() - 12, doc.internal.pageSize.getHeight() - 8, { align: "right" });
    }
    doc.save(`${fileSlug(schedule.title)}-${fileSlug(viewLabel)}-roadmap-poster.pdf`);
    return;
  }

  addSectionTitle(doc, "Executive Summary", 36);
  table(doc, autoTable, 42, ["Metric", "Value", "Metric", "Value"], summaryRows(schedule, items));
  addSectionTitle(doc, "Report Scope", 83);
  table(doc, autoTable, 89, ["Filter", "Value", "Filter", "Value"], filterRows(filters, viewLabel, dateWindowLabel, baselineNumber));

  doc.addPage();
  addHeader(doc, schedule, viewLabel);
  addSectionTitle(doc, "High Level Roadmap Timeline", 36);
  const timelineStart = selectedStart;
  const timelineLimit = 20;
  const timelineCandidateCount = timelineCandidates(items, timelineStart, selectedEnd).length;
  const timelineNote = `Timeline starts from selected date window (${formatDate(timelineStart.toISOString())}${selectedEnd ? ` to ${formatDate(selectedEnd.toISOString())}` : ""}); showing ${Math.min(timelineLimit, timelineCandidateCount)} of ${timelineCandidateCount} roadmap milestones.`;
  drawTimelineLegend(doc, 43, timelineNote);
  drawTimeline(doc, items, schedule, viewLabel, "High Level Roadmap Timeline", 78, timelineLimit, timelineStart, selectedEnd);

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
