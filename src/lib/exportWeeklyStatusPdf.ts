import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import type { TrackerChange, TrackerData, TrackerDecision, TrackerIssue, TrackerRisk } from "../types/reporting";
import { formatDate, parseDate } from "./dateUtils";

type DateWindow = {
  start?: Date;
  end?: Date;
  label: string;
};

type ExportWeeklyStatusOptions = {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
};

type Rgb = [number, number, number];
type AutoTable = typeof import("jspdf-autotable").default;
type TableRow = Array<string | number>;

const colours: Record<"ink" | "muted" | "deep" | "line" | "pale" | "green" | "amber" | "red" | "blue", Rgb> = {
  ink: [28, 38, 33],
  muted: [91, 105, 96],
  deep: [33, 76, 67],
  line: [199, 209, 203],
  pale: [243, 247, 245],
  green: [46, 125, 85],
  amber: [232, 117, 26],
  red: [179, 58, 50],
  blue: [61, 120, 169],
};

function fileSlug(value: string): string {
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "weekly-status";
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

function splitDigest(value?: string, limit = 3): string[] {
  const text = meaningfulText(value);
  if (!text) return [];
  const lineParts = text.split(/\n+/).map((part) => part.trim()).filter(Boolean);
  const parts = lineParts.length > 1 ? lineParts : text.split(/(?<=[.!?])\s+/).map((part) => part.trim()).filter(Boolean);
  return parts.slice(0, limit);
}

function bySoonest(a?: string, b?: string): number {
  return (parseDate(a)?.getTime() ?? Number.MAX_SAFE_INTEGER) - (parseDate(b)?.getTime() ?? Number.MAX_SAFE_INTEGER);
}

function dateWithin(value: string | undefined, window: DateWindow): boolean {
  if (!window.start && !window.end) return true;
  const date = parseDate(value);
  if (!date) return false;
  if (window.start && date < window.start) return false;
  if (window.end && date > window.end) return false;
  return true;
}

function weeklySummaryDate(summary: { meetingDate?: string; weekEnding?: string; lastUpdated?: string }): Date | undefined {
  return parseDate(summary.meetingDate) ?? parseDate(summary.weekEnding) ?? parseDate(summary.lastUpdated);
}

function sortedWeeklySummaries(tracker?: TrackerData) {
  return (tracker?.weeklySummaries ?? [])
    .slice()
    .sort((a, b) => (weeklySummaryDate(b)?.getTime() ?? 0) - (weeklySummaryDate(a)?.getTime() ?? 0));
}

function latestWeeklySummary(tracker?: TrackerData) {
  return sortedWeeklySummaries(tracker)[0];
}

function ragScore(value?: string): number | undefined {
  const text = normaliseText(value);
  if (text.includes("red")) return 3;
  if (text.includes("amber")) return 2;
  if (text.includes("green")) return 1;
  return undefined;
}

function weeklyMovement(tracker?: TrackerData): string | undefined {
  const summaries = sortedWeeklySummaries(tracker);
  const current = summaries[0];
  if (!current) return undefined;
  const explicit = meaningfulText(current.ragMovement);
  if (explicit) return explicit;
  const currentScore = ragScore(current.overallRag);
  const previousScore = ragScore(summaries[1]?.overallRag);
  if (!currentScore || !previousScore) return undefined;
  if (currentScore === previousScore) return "Stable";
  return currentScore > previousScore ? "Worsened" : "Improved";
}

function isRedOrAmber(value?: string): boolean {
  const label = normaliseText(value);
  return label.includes("red") || label.includes("amber") || label.includes("high");
}

function isOpenStatus(status?: string): boolean {
  const value = normaliseText(status);
  return !["complete", "completed", "closed", "done"].includes(value);
}

function itemImportance(item: ProgrammeItem): number {
  const level = normaliseText(item.milestoneLevel);
  if (item.executiveMilestone || level.includes("executive")) return 5;
  if (item.boardReportable || level.includes("board")) return 4;
  if (item.roadmapMilestone) return 3;
  if (item.governanceGate || item.decisionRequired) return 2;
  return 1;
}

function isHighLevelMilestone(item: ProgrammeItem): boolean {
  const level = normaliseText(item.milestoneLevel);
  return Boolean(item.executiveMilestone || item.boardReportable || item.roadmapMilestone || level.includes("executive") || level.includes("board"));
}

function programmeMilestones(schedule: ProgrammeSchedule): ProgrammeItem[] {
  return schedule.items
    .filter((item) => item.isMilestone || item.roadmapMilestone)
    .sort((a, b) => itemImportance(b) - itemImportance(a) || bySoonest(a.finishDate, b.finishDate));
}

function isOutstandingDecision(decision: TrackerDecision): boolean {
  const status = normaliseText(decision.status);
  if (["approved", "agreed", "decided", "closed", "complete", "completed", "done", "superseded", "cancelled", "not required"].includes(status)) {
    return false;
  }
  return Boolean(
    decision.dashboardFlag ||
      parseDate(decision.decisionRequiredBy) ||
      meaningfulText(decision.decisionRequiredByLabel) ||
      normaliseText(decision.decisionType).includes("required") ||
      normaliseText(decision.decisionType).includes("pending") ||
      status.includes("pending") ||
      status.includes("progress"),
  );
}

function decisionSort(a: TrackerDecision, b: TrackerDecision): number {
  return Number(Boolean(b.dashboardFlag)) - Number(Boolean(a.dashboardFlag)) || bySoonest(a.decisionRequiredBy ?? a.decisionDate, b.decisionRequiredBy ?? b.decisionDate);
}

function isSignificantChange(change: TrackerChange): boolean {
  const status = normaliseText(change.status);
  if (["closed", "complete", "completed", "done", "superseded", "cancelled"].includes(status)) return false;
  return Boolean(
    change.dashboardFlag ||
      meaningfulText(change.decisionRequired) ||
      meaningfulText(change.impactOnTime) ||
      meaningfulText(change.impactOnScope) ||
      meaningfulText(change.impactOnCost) ||
      meaningfulText(change.impactOnQualityOrBenefits),
  );
}

function changeSort(a: TrackerChange, b: TrackerChange): number {
  const aDate = parseDate(a.lastDiscussedDate ?? a.dateRaised);
  const bDate = parseDate(b.lastDiscussedDate ?? b.dateRaised);
  return Number(Boolean(b.dashboardFlag)) - Number(Boolean(a.dashboardFlag)) || (bDate?.getTime() ?? 0) - (aDate?.getTime() ?? 0);
}

function toneColour(value?: string): Rgb {
  const tone = normaliseText(value);
  if (tone.includes("red") || tone.includes("high")) return colours.red;
  if (tone.includes("amber") || tone.includes("medium")) return colours.amber;
  if (tone.includes("green") || tone.includes("low")) return colours.green;
  return colours.blue;
}

function setText(doc: JsPDF, colour: Rgb) {
  doc.setTextColor(colour[0], colour[1], colour[2]);
}

function addHeader(doc: JsPDF, title: string, reportDate: string) {
  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setFillColor(...colours.deep);
  doc.rect(0, 0, pageWidth, 28, "F");
  setText(doc, [255, 255, 255]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.text(title, 12, 11);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text("Weekly executive status", 12, 19);
  doc.text(`Report date ${formatDate(reportDate)}`, pageWidth - 12, 19, { align: "right" });
  setText(doc, colours.ink);
}

function ensureSpace(doc: JsPDF, y: number, requiredHeight: number): number {
  const pageHeight = doc.internal.pageSize.getHeight();
  if (y + requiredHeight <= pageHeight - 16) return y;
  doc.addPage();
  return 18;
}

function addBox(doc: JsPDF, x: number, y: number, w: number, h: number, title: string, body: string | string[]) {
  doc.setDrawColor(...colours.line);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(x, y, w, h, 2.2, 2.2, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  setText(doc, colours.muted);
  doc.text(title.toUpperCase(), x + 4, y + 6);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9.5);
  setText(doc, colours.ink);
  const lines = Array.isArray(body) ? body : doc.splitTextToSize(body, w - 8);
  doc.text(lines.slice(0, 7), x + 4, y + 13);
}

function textBoxHeight(doc: JsPDF, lines: string[], width: number): number {
  const wrapped = lines.flatMap((line) => doc.splitTextToSize(line, width));
  return Math.max(24, 14 + wrapped.length * 4.3);
}

function addNarrativeBox(doc: JsPDF, y: number, title: string, lines: string[], fallback: string): number {
  const margin = 12;
  const pageWidth = doc.internal.pageSize.getWidth();
  const width = pageWidth - margin * 2;
  const body = lines.length ? lines : [fallback];
  const height = textBoxHeight(doc, body, width - 12);
  y = ensureSpace(doc, y, height);
  doc.setDrawColor(...colours.line);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(margin, y, width, height, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  setText(doc, colours.ink);
  doc.text(title, margin + 5, y + 8);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  setText(doc, lines.length ? colours.ink : colours.muted);
  let cursor = y + 15;
  body.forEach((line) => {
    const wrapped = doc.splitTextToSize(line, width - 14);
    doc.text(lines.length ? "-" : "", margin + 5, cursor);
    doc.text(wrapped, margin + (lines.length ? 9 : 5), cursor);
    cursor += wrapped.length * 4.3;
  });
  return y + height + 5;
}

function sectionTitle(doc: JsPDF, title: string, y: number) {
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  setText(doc, colours.ink);
  doc.text(title, 12, y);
  doc.setDrawColor(...colours.line);
  doc.line(12, y + 3, doc.internal.pageSize.getWidth() - 12, y + 3);
}

function table(doc: JsPDF, autoTable: AutoTable, title: string, y: number, head: string[], body: TableRow[]): number {
  y = ensureSpace(doc, y, 32);
  sectionTitle(doc, title, y);
  autoTable(doc, {
    startY: y + 6,
    head: [head],
    body: body.length ? body : [["-", "No items currently flagged.", "-", "-"].slice(0, head.length)],
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
    margin: { left: 12, right: 12, top: 14, bottom: 16 },
    rowPageBreak: "avoid",
  });
  return ((doc as JsPDF & { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? y + 20) + 10;
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
    doc.text(`Page ${page} of ${pageCount}`, pageWidth - 12, pageHeight - 7, { align: "right" });
  }
}

export async function exportWeeklyStatusPdf({ schedule, tracker, dateWindow }: ExportWeeklyStatusOptions) {
  const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
  const weekly = latestWeeklySummary(tracker);
  const reportDate = weekly?.meetingDate ?? weekly?.weekEnding ?? new Date().toISOString();
  const forwardWindow = { ...dateWindow, start: dateWindow.start ?? parseDate(reportDate) };
  const upcomingMilestones = programmeMilestones(schedule)
    .filter((item) => isHighLevelMilestone(item) && dateWithin(item.finishDate, forwardWindow) && item.status !== "complete")
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate))
    .slice(0, 8);
  const risksIssues = [
    ...(tracker?.risks ?? [])
      .filter((risk) => isOpenStatus(risk.status) && (risk.dashboardFlag || isRedOrAmber(risk.rag)))
      .map((risk: TrackerRisk) => ({ id: `risk-${risk.id}`, title: risk.title, marker: risk.rag ?? risk.status ?? "Risk", owner: risk.owner ?? risk.stream ?? "-", update: risk.latestUpdate ?? risk.mitigation ?? risk.impact ?? "-" })),
    ...(tracker?.issues ?? [])
      .filter((issue) => isOpenStatus(issue.status) && (issue.dashboardFlag || isRedOrAmber(issue.rag) || isRedOrAmber(issue.priority)))
      .map((issue: TrackerIssue) => ({ id: `issue-${issue.id}`, title: issue.title, marker: issue.rag ?? issue.priority ?? issue.status ?? "Issue", owner: issue.owner ?? issue.stream ?? "-", update: issue.latestUpdate ?? issue.requiredAction ?? issue.impact ?? "-" })),
  ].slice(0, 8);
  const decisionsNeeded = (tracker?.decisions ?? []).filter(isOutstandingDecision).sort(decisionSort).slice(0, 8);
  const significantChanges = (tracker?.changes ?? []).filter(isSignificantChange).sort(changeSort).slice(0, 8);
  const movement = weeklyMovement(tracker) ?? "Not captured";
  const deliveryConfidence = meaningfulText(weekly?.goLiveConfidence) ?? "Not captured";
  const mainBlocker = meaningfulText(weekly?.mainBlocker) ?? risksIssues[0]?.title ?? "None flagged";
  const nextKeyDate = upcomingMilestones[0] ? `${formatDate(upcomingMilestones[0].finishDate)} - ${upcomingMilestones[0].name}` : "None in window";
  const progressItems = [...splitDigest(weekly?.keyProgress, 4), ...splitDigest(weekly?.whatChanged, 2)].slice(0, 5);
  const priorityItems = splitDigest(weekly?.priorityActions, 5);
  const leadershipAsk = meaningfulText(weekly?.askSteerNeeded) ?? meaningfulText(weekly?.decisionsNeeded) ?? decisionsNeeded[0]?.title ?? "No current leadership ask flagged.";

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  addHeader(doc, schedule.title, reportDate);

  let y = 36;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  setText(doc, colours.ink);
  doc.text("Status snapshot", 12, y);
  y += 5;
  doc.setDrawColor(...colours.line);
  doc.setFillColor(...colours.pale);
  doc.roundedRect(12, y, pageWidth - 24, 34, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  setText(doc, colours.ink);
  doc.text(doc.splitTextToSize(schedule.title, 120), 17, y + 9);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.2);
  setText(doc, colours.muted);
  const openingLine = meaningfulText(weekly?.openingLine) ?? meaningfulText(weekly?.ragRationale) ?? "Import the latest tracker to populate the weekly status update.";
  doc.text(doc.splitTextToSize(openingLine, 120).slice(0, 4), 17, y + 17);
  const rag = meaningfulText(weekly?.overallRag) ?? "Not captured";
  doc.setFillColor(...toneColour(rag));
  doc.roundedRect(pageWidth - 53, y + 5, 36, 20, 2.5, 2.5, "F");
  setText(doc, [255, 255, 255]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.text("OVERALL RAG", pageWidth - 35, y + 12, { align: "center" });
  doc.setFontSize(13);
  doc.text(rag, pageWidth - 35, y + 20, { align: "center" });
  y += 42;

  const cardWidth = (pageWidth - 30) / 2;
  addBox(doc, 12, y, cardWidth, 22, "Movement", movement);
  addBox(doc, 18 + cardWidth, y, cardWidth, 22, "Delivery confidence", deliveryConfidence);
  y += 27;
  addBox(doc, 12, y, cardWidth, 28, "Main blocker", mainBlocker);
  addBox(doc, 18 + cardWidth, y, cardWidth, 28, "Next key milestone", nextKeyDate);
  y += 34;

  y = addNarrativeBox(doc, y, "Last week", progressItems, "No weekly progress summary found.");
  y = addNarrativeBox(doc, y, "This week / next", priorityItems, "No priority actions summary found.");
  y = addNarrativeBox(doc, y, "Leadership ask", [leadershipAsk], "No current leadership ask flagged.");

  doc.addPage();
  y = 18;
  table(
    doc,
    autoTable,
    "Upcoming key milestones",
    y,
    ["Date", "Milestone", "Stream", "Status"],
    upcomingMilestones.map((item) => [formatDate(item.finishDate), item.name, item.stream ?? item.milestoneLevel ?? "-", item.status]),
  );
  y = ((doc as JsPDF & { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? y) + 10;
  y = table(
    doc,
    autoTable,
    "Risks / issues",
    y,
    ["Rating", "Item", "Owner / stream", "Latest update"],
    risksIssues.map((item) => [item.marker, item.title, item.owner, item.update]),
  );
  y = table(
    doc,
    autoTable,
    "Decisions needed",
    y,
    ["Date", "Decision", "Decision maker", "Status"],
    decisionsNeeded.map((decision) => [
      formatDate(decision.decisionRequiredBy ?? decision.decisionDate),
      decision.title,
      decision.decisionMaker ?? decision.owner ?? "-",
      decision.status ?? "Decision required",
    ]),
  );
  table(
    doc,
    autoTable,
    "Significant changes",
    y,
    ["Date", "Change", "Why it matters", "Owner"],
    significantChanges.map((change) => [
      formatDate(change.lastDiscussedDate ?? change.dateRaised),
      change.title,
      meaningfulText(change.decisionRequired) ??
        meaningfulText(change.impactOnTime) ??
        meaningfulText(change.impactOnScope) ??
        meaningfulText(change.impactOnCost) ??
        meaningfulText(change.impactOnQualityOrBenefits) ??
        change.latestUpdate ??
        "-",
      change.owner ?? change.decisionMaker ?? "-",
    ]),
  );

  addFooters(doc);
  doc.save(`${fileSlug(schedule.title)}-weekly-executive-status-a4.pdf`);
}
