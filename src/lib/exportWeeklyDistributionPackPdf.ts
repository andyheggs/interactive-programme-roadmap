import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import type { TrackerData, WeeklyStatusCuration, WeeklySummary } from "../types/reporting";
import { buildExecutiveRoadmapModel, executiveToneAssessment, executiveToneLabel, regularMilestoneToneAssessment } from "./executiveRoadmapData";
import { formatDate, parseDate } from "./dateUtils";

type DateWindow = {
  start?: Date;
  end?: Date;
  label: string;
};

type ExportWeeklyDistributionPackOptions = {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  curation?: WeeklyStatusCuration;
  contextItemUids?: string[];
  removedItemUids?: string[];
  laneOrderUids?: string[];
};

type Rgb = [number, number, number];
type AutoTable = typeof import("jspdf-autotable").default;
type TableRow = Array<string | number>;
type ExecutiveTone = ReturnType<typeof executiveToneAssessment>["tone"];

const colours: Record<"ink" | "muted" | "deep" | "line" | "pale" | "green" | "amber" | "red" | "blue" | "teal" | "purple" | "grey" | "white", Rgb> = {
  ink: [28, 38, 33],
  muted: [91, 105, 96],
  deep: [33, 76, 67],
  line: [199, 209, 203],
  pale: [243, 247, 245],
  green: [46, 125, 85],
  amber: [232, 117, 26],
  red: [179, 58, 50],
  blue: [61, 120, 169],
  teal: [20, 184, 166],
  purple: [123, 95, 196],
  grey: [126, 140, 132],
  white: [255, 255, 255],
};

const executiveToneColours: Record<ExecutiveTone, Rgb> = {
  green: colours.green,
  blue: colours.blue,
  teal: colours.teal,
  purple: colours.purple,
  amber: colours.amber,
  red: colours.red,
  grey: colours.grey,
};

const executiveToneFills: Record<ExecutiveTone, Rgb> = {
  green: [231, 245, 237],
  blue: [232, 242, 251],
  teal: [227, 250, 247],
  purple: [240, 236, 255],
  amber: [255, 240, 210],
  red: [255, 231, 229],
  grey: [237, 241, 239],
};

function fileSlug(value: string): string {
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "weekly-pack";
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

function weeklyProgrammeTitle(value?: string): string {
  const title = (value ?? "").trim() || "Programme Delivery";
  const withoutVersion = title
    .replace(/\s+(?:-|\u2013)\s*v\d.*$/i, "")
    .replace(/\s*\/\s*/g, " ")
    .trim();
  if (/^daf programme delivery$/i.test(withoutVersion)) return "Data Asset Foundation Programme Delivery";
  if (/^daf\b/i.test(withoutVersion)) return withoutVersion.replace(/^daf\b/i, "Data Asset Foundation");
  return withoutVersion || title;
}

function splitDigest(value?: string, limit = 5): string[] {
  const text = meaningfulText(value);
  if (!text) return [];
  const lineParts = text.split(/\n+/).map((part) => part.trim()).filter(Boolean);
  const parts = lineParts.length > 1 ? lineParts : text.split(/(?<=[.!?])\s+/).map((part) => part.trim()).filter(Boolean);
  return parts.slice(0, limit);
}

function bySoonest(a?: string, b?: string): number {
  return (parseDate(a)?.getTime() ?? Number.MAX_SAFE_INTEGER) - (parseDate(b)?.getTime() ?? Number.MAX_SAFE_INTEGER);
}

function weeklySummaryDate(summary: { meetingDate?: string; weekEnding?: string; lastUpdated?: string }): Date | undefined {
  return parseDate(summary.meetingDate) ?? parseDate(summary.weekEnding) ?? parseDate(summary.lastUpdated);
}

function latestWeeklySummary(tracker?: TrackerData): WeeklySummary | undefined {
  return (tracker?.weeklySummaries ?? [])
    .slice()
    .sort((a, b) => (weeklySummaryDate(b)?.getTime() ?? 0) - (weeklySummaryDate(a)?.getTime() ?? 0))[0];
}

function isOpenStatus(status?: string): boolean {
  const value = normaliseText(status);
  return !["complete", "completed", "closed", "done", "resolved", "implemented"].includes(value);
}

function isCompleteStatus(status?: string): boolean {
  return !isOpenStatus(status);
}

function isRedOrAmber(value?: string): boolean {
  const label = normaliseText(value);
  return label.includes("red") || label.includes("amber") || label.includes("high");
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

function programmeDeliveryOutcome(schedule: ProgrammeSchedule, outcomes: ProgrammeItem[]): ProgrammeItem | undefined {
  const candidates = outcomes.length ? outcomes : executiveMilestoneItems(schedule);
  return candidates.find((item) => /platform.*go live|go live/i.test(item.name))
    ?? candidates.slice().sort((a, b) => bySoonest(b.finishDate, a.finishDate))[0]
    ?? schedule.items.slice().sort((a, b) => bySoonest(b.finishDate, a.finishDate))[0];
}

function forecastToGoLiveLabel(schedule: ProgrammeSchedule): string {
  const outcome = programmeDeliveryOutcome(schedule, executiveMilestoneItems(schedule));
  if (outcome?.finishDate) return `${formatDate(outcome.finishDate)} - ${outcome.name}`;
  return formatDate(schedule.finishDate);
}

function generatedStatusSummary(weekly?: WeeklySummary, mainBlocker?: string): string | undefined {
  const progress = [...splitDigest(weekly?.keyProgress, 2), ...splitDigest(weekly?.whatChanged, 1)].slice(0, 2);
  const priorities = splitDigest(weekly?.priorityActions, 2);
  const blocker = meaningfulText(mainBlocker) ?? meaningfulText(weekly?.keyRisksOrIssues) ?? meaningfulText(weekly?.ragRationale);
  const rag = meaningfulText(weekly?.overallRag);
  const parts: string[] = [];
  if (rag) parts.push(`The programme is currently ${rag}.`);
  if (progress.length) parts.push(progress.join(" "));
  if (priorities.length) parts.push(`Immediate focus: ${priorities.join(" ")}`);
  if (blocker) parts.push(`Main blocker: ${blocker}.`);
  return parts.length ? parts.join(" ") : undefined;
}

function toneColour(value?: string): Rgb {
  const text = normaliseText(value);
  if (text.includes("red") || text.includes("overdue") || text.includes("blocked")) return colours.red;
  if (text.includes("amber") || text.includes("soon") || text.includes("high")) return colours.amber;
  if (text.includes("green") || text.includes("complete")) return colours.green;
  return colours.blue;
}

function addHeader(doc: JsPDF, title: string, subtitle: string, reportDate: string) {
  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setFillColor(...colours.deep);
  doc.rect(0, 0, pageWidth, 24, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text(title, 12, 9);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(subtitle, 12, 17);
  doc.text(`Generated ${formatDate(reportDate)}`, pageWidth - 12, 17, { align: "right" });
}

function addFooter(doc: JsPDF) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const pageCount = doc.getNumberOfPages();
  for (let page = 1; page <= pageCount; page += 1) {
    doc.setPage(page);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    doc.setTextColor(...colours.muted);
    doc.text(`Page ${page} of ${pageCount}`, pageWidth - 12, pageHeight - 6, { align: "right" });
  }
}

function sectionTitle(doc: JsPDF, title: string, y: number): number {
  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.setTextColor(...colours.ink);
  doc.text(title, 12, y);
  doc.setDrawColor(...colours.line);
  doc.line(12, y + 3, doc.internal.pageSize.getWidth() - 12, y + 3);
  return y + 10;
}

function addTextBox(doc: JsPDF, title: string, body: string | string[], x: number, y: number, w: number, minHeight = 30): number {
  const lines = Array.isArray(body) ? body : doc.splitTextToSize(body, w - 10);
  const h = Math.max(minHeight, 16 + lines.length * 4.4);
  doc.setDrawColor(...colours.line);
  doc.setFillColor(...colours.pale);
  doc.roundedRect(x, y, w, h, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.setTextColor(...colours.muted);
  doc.text(title.toUpperCase(), x + 5, y + 7);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.4);
  doc.setTextColor(...colours.ink);
  doc.text(lines, x + 5, y + 15);
  return h;
}

function setRgbFill(doc: JsPDF, colour: Rgb) {
  doc.setFillColor(colour[0], colour[1], colour[2]);
}

function setRgbDraw(doc: JsPDF, colour: Rgb) {
  doc.setDrawColor(colour[0], colour[1], colour[2]);
}

function setRgbText(doc: JsPDF, colour: Rgb) {
  doc.setTextColor(colour[0], colour[1], colour[2]);
}

function drawRoadmapLegend(doc: JsPDF, x: number, y: number, w: number): number {
  const items: Array<{ tone: ExecutiveTone; label: string }> = [
    { tone: "green", label: "Complete" },
    { tone: "blue", label: "Ongoing" },
    { tone: "teal", label: "Future" },
    { tone: "amber", label: "Date assumption" },
    { tone: "red", label: "Late" },
  ];
  setRgbDraw(doc, colours.line);
  setRgbFill(doc, colours.pale);
  doc.roundedRect(x, y, w, 17, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7);
  setRgbText(doc, colours.muted);
  doc.text("Colour status", x + 4, y + 6.5);
  let cursor = x + 34;
  items.forEach((item) => {
    setRgbFill(doc, executiveToneColours[item.tone]);
    doc.circle(cursor, y + 6, 2.4, "F");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(6.9);
    setRgbText(doc, colours.ink);
    doc.text(item.label, cursor + 5, y + 8);
    cursor += doc.getTextWidth(item.label) + 16;
  });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(6.6);
  setRgbText(doc, colours.muted);
  doc.text("Lane milestones use Project Status plus Date Assumption. Executive cards retain the executive milestone RAG rule.", x + 4, y + 14);
  return y + 24;
}

function drawOutcomeCard(doc: JsPDF, item: ProgrammeItem, x: number, y: number, w: number, h: number) {
  const assessment = executiveToneAssessment(item);
  const tone = assessment.tone;
  setRgbDraw(doc, executiveToneColours[tone]);
  setRgbFill(doc, executiveToneFills[tone]);
  doc.roundedRect(x, y, w, h, 2.5, 2.5, "FD");
  setRgbFill(doc, executiveToneColours[tone]);
  doc.rect(x, y, w, 4, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.2);
  setRgbText(doc, colours.muted);
  doc.text(executiveToneLabel(item), x + 4, y + 11);
  doc.setFontSize(8.2);
  setRgbText(doc, colours.ink);
  doc.text(doc.splitTextToSize(item.name, w - 8).slice(0, 3), x + 4, y + 20);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  setRgbText(doc, colours.muted);
  doc.text(formatDate(item.finishDate), x + 4, y + h - 7);
}

function drawRoadmapLane(doc: JsPDF, path: { outcome: ProgrammeItem; dependencies: ProgrammeItem[] }, x: number, y: number, w: number, h: number) {
  const labelW = 48;
  const outcomeW = 58;
  const gap = 6;
  const pathX = x + labelW + gap;
  const outcomeX = x + w - outcomeW;
  const pathW = outcomeX - pathX - gap;
  const centreY = y + h / 2 - 1;
  const dependencies = path.dependencies.slice(0, 6);

  setRgbDraw(doc, colours.line);
  setRgbFill(doc, colours.white);
  doc.roundedRect(x, y, w, h, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8.4);
  setRgbText(doc, colours.ink);
  doc.text(doc.splitTextToSize(path.outcome.name, labelW - 6).slice(0, 3), x + 4, y + 10);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(6.5);
  setRgbText(doc, colours.muted);
  doc.text("Executive milestone", x + 4, y + h - 6);

  setRgbDraw(doc, colours.line);
  doc.setLineWidth(0.55);
  doc.line(pathX, centreY, outcomeX - 5, centreY);
  doc.line(outcomeX - 5, centreY, outcomeX - 9, centreY - 3);
  doc.line(outcomeX - 5, centreY, outcomeX - 9, centreY + 3);

  if (!dependencies.length) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.2);
    setRgbText(doc, colours.muted);
    doc.text("No linked predecessor milestones shown.", pathX + 4, centreY - 4);
  } else {
    const step = dependencies.length > 1 ? pathW / (dependencies.length - 1) : 0;
    const textW = Math.min(34, Math.max(24, pathW / Math.max(1, dependencies.length) - 3));
    dependencies.forEach((item, index) => {
      const dotX = dependencies.length === 1 ? pathX + pathW / 2 : pathX + index * step;
      const assessment = regularMilestoneToneAssessment(item);
      const tone = assessment.tone;
      setRgbFill(doc, executiveToneColours[tone]);
      doc.circle(dotX, centreY, 2.8, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(6.4);
      setRgbText(doc, colours.muted);
      doc.text(formatDate(item.finishDate).replace(" 20", " "), dotX, y + 10, { align: "center" });
      doc.setFontSize(6.7);
      setRgbText(doc, colours.ink);
      doc.text(doc.splitTextToSize(item.name, textW).slice(0, 3), dotX, centreY + 8, { align: "center" });
    });
  }

  drawOutcomeCard(doc, path.outcome, outcomeX, y + 5, outcomeW, h - 10);
}

function addExecutiveRoadmapVisualPages(
  doc: JsPDF,
  schedule: ProgrammeSchedule,
  reportDate: string,
  model: ReturnType<typeof buildExecutiveRoadmapModel>,
) {
  const title = weeklyProgrammeTitle(schedule.title);
  const paths = model.paths.length ? model.paths : [];
  const lanesPerPage = 2;
  const chunks = paths.length ? Array.from({ length: Math.ceil(paths.length / lanesPerPage) }, (_, index) => paths.slice(index * lanesPerPage, index * lanesPerPage + lanesPerPage)) : [[]];

  chunks.forEach((chunk, index) => {
    doc.addPage("a4", "landscape");
    addHeader(doc, title, "Executive roadmap visual", reportDate);
    let y = sectionTitle(doc, index ? `Executive roadmap visual ${index + 1}` : "Executive roadmap visual", 36);
    const pageWidth = doc.internal.pageSize.getWidth();
    y = drawRoadmapLegend(doc, 12, y, pageWidth - 24);
    if (!chunk.length) {
      addTextBox(doc, "No executive milestones found", "Flag the high-level outcomes in Microsoft Project using the Executive Milestones field, then re-import the XML.", 12, y, pageWidth - 24, 30);
      return;
    }
    chunk.forEach((path) => {
      drawRoadmapLane(doc, path, 12, y, pageWidth - 24, 54);
      y += 62;
    });
  });
}

function table(
  doc: JsPDF,
  autoTable: AutoTable,
  y: number,
  head: string[],
  body: TableRow[],
  columnStyles: Record<number, object> = {},
  headColour: Rgb = colours.deep,
): number {
  autoTable(doc, {
    startY: y,
    margin: { left: 12, right: 12, bottom: 14 },
    head: [head],
    body,
    theme: "grid",
    showHead: "everyPage",
    styles: { font: "helvetica", fontSize: 7.5, cellPadding: 2.2, textColor: colours.ink, lineColor: colours.line, lineWidth: 0.1, overflow: "linebreak", valign: "top" },
    headStyles: { fillColor: headColour, textColor: [255, 255, 255], fontStyle: "bold", fontSize: 7.4 },
    alternateRowStyles: { fillColor: colours.pale },
    columnStyles,
  });
  return ((doc as JsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 8;
}

export async function exportWeeklyDistributionPackPdf({
  schedule,
  tracker,
  dateWindow,
  curation,
  contextItemUids,
  removedItemUids,
  laneOrderUids,
}: ExportWeeklyDistributionPackOptions) {
  const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const tablePlugin = autoTable as AutoTable;
  const reportDate = new Date().toISOString();
  const title = weeklyProgrammeTitle(schedule.title);
  const weekly = latestWeeklySummary(tracker);
  const openRisks = (tracker?.risks ?? []).filter((item) => isOpenStatus(item.status));
  const openIssues = (tracker?.issues ?? []).filter((item) => isOpenStatus(item.status));
  const riskIssues = [
    ...openRisks.map((item) => ({ kind: "Risk", title: item.title, owner: item.owner ?? item.stream ?? "Not set", marker: item.rag ?? item.likelihood ?? "Risk", update: item.latestUpdate ?? item.mitigation ?? item.impact ?? "" })),
    ...openIssues.map((item) => ({ kind: "Issue", title: item.title, owner: item.owner ?? item.stream ?? "Not set", marker: item.rag ?? item.priority ?? "Issue", update: item.latestUpdate ?? item.requiredAction ?? item.impact ?? "" })),
  ].sort((a, b) => Number(isRedOrAmber(b.marker)) - Number(isRedOrAmber(a.marker)) || a.title.localeCompare(b.title));
  const mainBlocker = meaningfulText(weekly?.mainBlocker) ?? riskIssues[0]?.title ?? "None flagged";
  const statusSummary =
    curation?.statusSummaryOverride ??
    meaningfulText(weekly?.executiveStatusSummary) ??
    meaningfulText(weekly?.openingLine) ??
    meaningfulText(weekly?.ragRationale) ??
    generatedStatusSummary(weekly, mainBlocker) ??
    "Import the latest tracker to populate the weekly status update.";
  const rag = meaningfulText(weekly?.overallRag) ?? "Not captured";
  const model = buildExecutiveRoadmapModel(schedule, tracker, dateWindow, { contextItemUids, removedItemUids, laneOrderUids });
  const nextMilestone = programmeMilestones(schedule)
    .filter((item) => isOpenStatus(item.status) && parseDate(item.finishDate) && (!dateWindow.start || parseDate(item.finishDate)! >= dateWindow.start) && (!dateWindow.end || parseDate(item.finishDate)! <= dateWindow.end))
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate))[0];
  const progressItems = splitDigest(weekly?.keyProgress, 5);
  const priorityItems = splitDigest(weekly?.priorityActions, 5);
  const whatChangedItems = splitDigest(curation?.whatChangedOverride ?? weekly?.whatChanged, 5);

  addHeader(doc, title, "Weekly distribution pack", reportDate);
  let y = sectionTitle(doc, "Email-ready summary", 36);
  const pageWidth = doc.internal.pageSize.getWidth();
  const ragBoxWidth = 38;
  const ragGap = 10;
  const ragBoxX = pageWidth - 12 - ragBoxWidth;
  const textWidth = ragBoxX - 12 - ragGap;
  const summaryHeight = addTextBox(doc, title, statusSummary, 12, y, textWidth, 44);
  doc.setFillColor(...toneColour(rag));
  doc.roundedRect(ragBoxX, y, ragBoxWidth, summaryHeight, 2.5, 2.5, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.text("OVERALL RAG", ragBoxX + ragBoxWidth / 2, y + 12, { align: "center" });
  doc.setFontSize(16);
  doc.text(rag, ragBoxX + ragBoxWidth / 2, y + 25, { align: "center" });
  doc.setFontSize(7);
  doc.text(formatDate(reportDate), ragBoxX + ragBoxWidth / 2, y + 35, { align: "center" });
  y += summaryHeight + 8;

  doc.addPage();
  addHeader(doc, title, "Weekly executive status", reportDate);
  y = sectionTitle(doc, "Weekly executive status", 36);
  const cardWidth = (pageWidth - 30) / 2;
  const left = 12;
  const right = 18 + cardWidth;
  addTextBox(doc, "Delivery confidence", meaningfulText(weekly?.goLiveConfidence) ?? "Not captured", left, y, cardWidth, 24);
  addTextBox(doc, "Forecast to go live", forecastToGoLiveLabel(schedule), right, y, cardWidth, 24);
  y += 32;
  addTextBox(doc, "Main blocker", mainBlocker, left, y, cardWidth, 28);
  addTextBox(doc, "Next milestone", nextMilestone ? `${formatDate(nextMilestone.finishDate)} - ${nextMilestone.name}` : "None in selected window", right, y, cardWidth, 28);
  y += 38;
  y += addTextBox(doc, "Last week", progressItems.length ? progressItems : ["No weekly progress summary found."], left, y, pageWidth - 24, 30) + 7;
  y += addTextBox(doc, "This week / next", priorityItems.length ? priorityItems : ["No priority actions summary found."], left, y, pageWidth - 24, 30) + 7;
  addTextBox(doc, "What changed this week", whatChangedItems.length ? whatChangedItems : ["No material changes captured in the latest weekly row."], left, y, pageWidth - 24, 30);

  addExecutiveRoadmapVisualPages(doc, schedule, reportDate, model);

  addFooter(doc);
  doc.save(`${fileSlug(schedule.title)}-weekly-distribution-pack.pdf`);
}
