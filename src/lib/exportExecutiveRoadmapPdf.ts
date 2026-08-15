import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import type { TrackerData } from "../types/reporting";
import { formatDate, parseDate } from "./dateUtils";
import { executiveToneAssessment as assessExecutiveTone, executiveToneLabel } from "./executiveRoadmapData";

type DateWindow = {
  start?: Date;
  end?: Date;
  label: string;
};

type ExportExecutiveRoadmapOptions = {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
};

type ExecutiveTone = "green" | "blue" | "amber" | "red" | "grey";
type Rgb = [number, number, number];
type AutoTable = typeof import("jspdf-autotable").default;
type TableRow = Array<string | number>;

type ExecutivePathNode = {
  item: ProgrammeItem;
  depth: number;
};

const colours: Record<"ink" | "muted" | "deep" | "line" | "pale" | "green" | "amber" | "red" | "blue" | "grey", Rgb> = {
  ink: [28, 38, 33],
  muted: [91, 105, 96],
  deep: [33, 76, 67],
  line: [199, 209, 203],
  pale: [243, 247, 245],
  green: [46, 125, 85],
  amber: [255, 138, 0],
  red: [179, 58, 50],
  blue: [61, 120, 169],
  grey: [126, 140, 132],
};

function fileSlug(value: string): string {
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "executive-roadmap";
}

function normaliseText(value?: string): string {
  return (value ?? "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function meaningfulText(value?: string): string | undefined {
  const text = value?.trim();
  if (!text) return undefined;
  const normalised = normaliseText(text);
  if (
    !normalised ||
    normalised === "na" ||
    normalised === "n a" ||
    normalised === "none" ||
    normalised === "nil" ||
    normalised === "not set" ||
    normalised === "not captured" ||
    normalised.startsWith("not stated") ||
    normalised.startsWith("to be confirmed") ||
    normalised === "tbc"
  ) {
    return undefined;
  }
  return text;
}

function bySoonest(a?: string, b?: string): number {
  return (parseDate(a)?.getTime() ?? Number.MAX_SAFE_INTEGER) - (parseDate(b)?.getTime() ?? Number.MAX_SAFE_INTEGER);
}

function itemImportance(item: ProgrammeItem): number {
  const level = normaliseText(item.milestoneLevel);
  if (item.executiveMilestone || level.includes("executive")) return 5;
  if (item.boardReportable || level.includes("board")) return 4;
  if (item.roadmapMilestone) return 3;
  if (item.governanceGate || item.decisionRequired) return 2;
  return 1;
}

function programmeMilestones(schedule: ProgrammeSchedule): ProgrammeItem[] {
  return schedule.items
    .filter((item) => item.isMilestone || item.roadmapMilestone)
    .sort((a, b) => itemImportance(b) - itemImportance(a) || bySoonest(a.finishDate, b.finishDate));
}

function executiveMilestoneItems(schedule: ProgrammeSchedule): ProgrammeItem[] {
  const executive = schedule.items
    .filter((item) => item.executiveMilestone || normaliseText(item.milestoneLevel) === "executive milestone")
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
  if (executive.length) return executive;
  return programmeMilestones(schedule).filter((item) => itemImportance(item) >= 4).slice(0, 5);
}

function isDeliveredItem(item: ProgrammeItem): boolean {
  return item.status === "complete" || item.percentComplete === 100;
}

function isHistoricDeliveredItem(item: ProgrammeItem, windowStart?: Date): boolean {
  const finishDate = parseDate(item.finishDate);
  return Boolean(isDeliveredItem(item) && finishDate && windowStart && finishDate < windowStart);
}

function programmeDeliveryOutcome(schedule: ProgrammeSchedule, outcomes: ProgrammeItem[]): ProgrammeItem | undefined {
  const candidates = outcomes.length ? outcomes : executiveMilestoneItems(schedule);
  return candidates.find((item) => /platform.*go live|go live/i.test(item.name))
    ?? candidates.slice().sort((a, b) => bySoonest(b.finishDate, a.finishDate))[0]
    ?? schedule.items.slice().sort((a, b) => bySoonest(b.finishDate, a.finishDate))[0];
}

function executiveDependencyScore(item: ProgrammeItem): number {
  const level = `${item.milestoneLevel ?? ""} ${item.dependencyLevel ?? ""}`.toLowerCase();
  let score = 0;
  if (item.executiveMilestone) score += 100;
  if (item.boardReportable) score += 60;
  if (item.roadmapMilestone) score += 50;
  if (item.decisionRequired) score += 35;
  if (item.governanceGate) score += 25;
  if (level.includes("executive")) score += 80;
  if (level.includes("board")) score += 45;
  if (level.includes("gate")) score += 30;
  return score;
}

function executivePathRelevance(node: ExecutivePathNode): number {
  const item = node.item;
  const level = `${item.milestoneLevel ?? ""} ${item.dependencyLevel ?? ""}`.toLowerCase();
  const name = normaliseText(item.name);
  let score = node.depth === 1 ? 120 : Math.max(0, 72 - node.depth * 8);
  if (item.executiveMilestone) score += 100;
  if (item.boardReportable) score += 70;
  if (item.roadmapMilestone) score += 55;
  if (level.includes("level 1")) score += 65;
  if (level.includes("level 2")) score += 45;
  if (level.includes("executive")) score += 60;
  if (level.includes("board")) score += 45;
  if (level.includes("roadmap")) score += 35;
  [
    "approved",
    "approval",
    "consultation period closed",
    "consultation response",
    "contract signed",
    "contract negotiations",
    "companies registry",
    "registrar employed",
    "platform ready",
    "readiness",
    "go live",
    "ready for approval",
    "submitted to dfe board",
  ].forEach((keyword) => {
    if (name.includes(keyword)) score += 28;
  });
  if (item.isSummary) score -= 120;
  if (!item.isActive) score -= 120;
  return score;
}

function visibleExecutivePathItem(item: ProgrammeItem): boolean {
  return Boolean(item.isActive && !item.isSummary && (item.isMilestone || meaningfulText(item.milestoneLevel)));
}

function collectPredecessorChainWithDepth(outcome: ProgrammeItem, byUid: Map<string, ProgrammeItem>): ExecutivePathNode[] {
  const found = new Map<string, ExecutivePathNode>();
  const seen = new Set<string>();
  const walk = (item: ProgrammeItem, depth: number) => {
    if (depth > 8 || seen.has(item.uid)) return;
    seen.add(item.uid);
    item.predecessors.forEach((link) => {
      if (!link.predecessorUid) return;
      const predecessor = byUid.get(link.predecessorUid);
      if (!predecessor) return;
      const existing = found.get(predecessor.uid);
      if (!existing || depth + 1 < existing.depth) found.set(predecessor.uid, { item: predecessor, depth: depth + 1 });
      walk(predecessor, depth + 1);
    });
  };
  walk(outcome, 0);
  return [...found.values()];
}

function directPredecessors(item: ProgrammeItem, byUid: Map<string, ProgrammeItem>): ProgrammeItem[] {
  const predecessors = item.predecessors
    .map((link) => (link.predecessorUid ? byUid.get(link.predecessorUid) : undefined))
    .filter((predecessor): predecessor is ProgrammeItem => Boolean(predecessor?.isActive && !predecessor.isSummary));
  return [...new Map(predecessors.map((predecessor) => [predecessor.uid, predecessor])).values()]
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
}

function closestDirectPredecessors(item: ProgrammeItem, byUid: Map<string, ProgrammeItem>, limit = 2): ProgrammeItem[] {
  const targetDate = parseDate(item.startDate) ?? parseDate(item.finishDate);
  return directPredecessors(item, byUid)
    .sort((a, b) => {
      const aDate = parseDate(a.finishDate ?? a.startDate);
      const bDate = parseDate(b.finishDate ?? b.startDate);
      if (targetDate && aDate && bDate) {
        return Math.abs(targetDate.getTime() - aDate.getTime()) - Math.abs(targetDate.getTime() - bDate.getTime());
      }
      return (bDate?.getTime() ?? 0) - (aDate?.getTime() ?? 0);
    })
    .slice(0, limit)
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
}

function collectPredecessorDependencies(outcome: ProgrammeItem, byUid: Map<string, ProgrammeItem>, windowStart?: Date): ProgrammeItem[] {
  const chain = collectPredecessorChainWithDepth(outcome, byUid);
  return chain
    .filter((node) => visibleExecutivePathItem(node.item))
    .filter((node) => !isHistoricDeliveredItem(node.item, windowStart))
    .sort((a, b) => bySoonest(a.item.finishDate, b.item.finishDate) || a.depth - b.depth)
    .map((node) => node.item);
}

function executiveTone(item?: ProgrammeItem): ExecutiveTone {
  return assessExecutiveTone(item).tone as ExecutiveTone;
}

function weeklySummaryDate(summary: { meetingDate?: string; weekEnding?: string; lastUpdated?: string }): Date | undefined {
  return parseDate(summary.meetingDate) ?? parseDate(summary.weekEnding) ?? parseDate(summary.lastUpdated);
}

function latestWeeklySummary(tracker?: TrackerData) {
  return (tracker?.weeklySummaries ?? [])
    .slice()
    .sort((a, b) => (weeklySummaryDate(b)?.getTime() ?? 0) - (weeklySummaryDate(a)?.getTime() ?? 0))[0];
}

function setText(doc: JsPDF, colour: Rgb) {
  doc.setTextColor(colour[0], colour[1], colour[2]);
}

function addHeader(doc: JsPDF, schedule: ProgrammeSchedule, reportDate: string, dateWindow: DateWindow) {
  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setFillColor(...colours.deep);
  doc.rect(0, 0, pageWidth, 25, "F");
  setText(doc, [255, 255, 255]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.text("DAF Executive Delivery Roadmap", 12, 10);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(schedule.title, 12, 18);
  doc.text(`Reporting date ${formatDate(reportDate)}`, pageWidth - 12, 10, { align: "right" });
  doc.text(`Date window: ${dateWindow.label}`, pageWidth - 12, 18, { align: "right" });
}

function addFooters(doc: JsPDF) {
  const pageCount = doc.getNumberOfPages();
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  for (let page = 1; page <= pageCount; page += 1) {
    doc.setPage(page);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    setText(doc, colours.muted);
    doc.text(`Page ${page} of ${pageCount}`, pageWidth - 12, pageHeight - 6, { align: "right" });
  }
}

function sectionTitle(doc: JsPDF, title: string, y: number) {
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  setText(doc, colours.ink);
  doc.text(title, 12, y);
  doc.setDrawColor(...colours.line);
  doc.line(12, y + 3, doc.internal.pageSize.getWidth() - 12, y + 3);
}

function table(
  doc: JsPDF,
  autoTable: AutoTable,
  schedule: ProgrammeSchedule,
  reportDate: string,
  dateWindow: DateWindow,
  title: string,
  y: number,
  head: string[],
  body: TableRow[],
  columnStyles: Record<number, object> = {},
): number {
  const pageHeight = doc.internal.pageSize.getHeight();
  if (y > pageHeight - 35) {
    doc.addPage();
    y = 34;
  }
  sectionTitle(doc, title, y);
  autoTable(doc, {
    startY: y + 6,
    head: [head],
    body: body.length ? body : [["-", "No linked items found.", "-", "-", "-"].slice(0, head.length)],
    theme: "grid",
    showHead: "everyPage",
    styles: {
      font: "helvetica",
      fontSize: 7.5,
      cellPadding: 2.1,
      overflow: "linebreak",
      textColor: colours.ink,
      lineColor: colours.line,
      lineWidth: 0.1,
      valign: "top",
    },
    headStyles: {
      fillColor: colours.deep,
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    alternateRowStyles: {
      fillColor: colours.pale,
    },
    margin: { left: 12, right: 12, top: 30, bottom: 14 },
    rowPageBreak: "avoid",
    columnStyles,
    didParseCell: (data) => {
      if (data.section !== "body") return;
      if (head.includes("RAG") && data.column.index === head.indexOf("RAG")) {
        const tone = normaliseText(String(data.cell.raw));
        if (tone.includes("red")) data.cell.styles.textColor = colours.red;
        else if (tone.includes("amber")) data.cell.styles.textColor = colours.amber;
        else if (tone.includes("green")) data.cell.styles.textColor = colours.green;
        else data.cell.styles.textColor = colours.blue;
        data.cell.styles.fontStyle = "bold";
      }
      if (head.includes("Status") && data.column.index === head.indexOf("Status")) {
        data.cell.styles.fontStyle = "bold";
      }
    },
    didDrawPage: () => addHeader(doc, schedule, reportDate, dateWindow),
  });
  return ((doc as JsPDF & { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? y + 20) + 10;
}

export async function exportExecutiveRoadmapPdf({ schedule, tracker, dateWindow }: ExportExecutiveRoadmapOptions) {
  const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
  const reportDate = new Date().toISOString();
  const windowStart = dateWindow.start ?? parseDate(reportDate);
  const outcomesAll = executiveMilestoneItems(schedule);
  const outcomes = outcomesAll.filter((item) => !isHistoricDeliveredItem(item, windowStart));
  const deliveryOutcome = programmeDeliveryOutcome(schedule, outcomesAll);
  const weekly = latestWeeklySummary(tracker);
  const programmeStatus = meaningfulText(weekly?.overallRag) ?? "Not captured";
  const byUid = new Map(schedule.items.map((item) => [item.uid, item]));
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

  addHeader(doc, schedule, reportDate, dateWindow);
  let y = 34;

  autoTable(doc, {
    startY: y,
    head: [["Programme status", "Original delivery plan", "Current forecast", "Forecast basis"]],
    body: [[
      programmeStatus,
      formatDate(deliveryOutcome?.baselineFinish),
      formatDate(deliveryOutcome?.finishDate ?? schedule.finishDate),
      deliveryOutcome?.name ?? "Programme finish",
    ]],
    theme: "grid",
    styles: { font: "helvetica", fontSize: 8.5, cellPadding: 3, textColor: colours.ink, lineColor: colours.line, lineWidth: 0.1 },
    headStyles: { fillColor: colours.pale, textColor: colours.muted, fontStyle: "bold" },
    bodyStyles: { fontStyle: "bold" },
    margin: { left: 12, right: 12, bottom: 14 },
    columnStyles: {
      0: { cellWidth: 48 },
      1: { cellWidth: 48 },
      2: { cellWidth: 48 },
      3: { cellWidth: 125 },
    },
  });
  y = ((doc as JsPDF & { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? y) + 10;

  y = table(
    doc,
    autoTable,
    schedule,
    reportDate,
    dateWindow,
    "Executive milestones",
    y,
    ["Date", "RAG", "RAG reason", "Executive milestone", "Baseline", "Stream"],
    outcomes.map((item) => [
      formatDate(item.finishDate),
      executiveToneLabel(item),
      assessExecutiveTone(item).summary,
      item.name,
      formatDate(item.baselineFinish),
      item.stream ?? item.milestoneLevel ?? "-",
    ]),
    {
      0: { cellWidth: 24, fontStyle: "bold" },
      1: { cellWidth: 20, fontStyle: "bold" },
      2: { cellWidth: 55 },
      3: { cellWidth: 95, fontStyle: "bold" },
      4: { cellWidth: 25 },
      5: { cellWidth: 50 },
    },
  );

  outcomes.forEach((outcome) => {
    const dependencies = collectPredecessorDependencies(outcome, byUid, windowStart);
    const rows = [...dependencies, outcome].map((item) => {
      const direct = closestDirectPredecessors(item, byUid, 2)
        .map((predecessor) => predecessor.name)
        .join("; ");
      return [
        formatDate(item.finishDate),
        item.uid === outcome.uid ? "Executive outcome" : "Predecessor",
        executiveToneLabel(item),
        assessExecutiveTone(item).summary,
        item.name,
        direct || "-",
      ];
    });
    y = table(
      doc,
      autoTable,
      schedule,
      reportDate,
      dateWindow,
      outcome.targetMilestone || outcome.name,
      y,
      ["Date", "Type", "Status", "RAG reason", "Item", "Immediate predecessors"],
      rows,
      {
        0: { cellWidth: 24, fontStyle: "bold" },
        1: { cellWidth: 28 },
        2: { cellWidth: 22, fontStyle: "bold" },
        3: { cellWidth: 52 },
        4: { cellWidth: 75, fontStyle: "bold" },
        5: { cellWidth: 68 },
      },
    );
  });

  addFooters(doc);
  doc.save(`${fileSlug(schedule.title)}-executive-roadmap-a4.pdf`);
}
