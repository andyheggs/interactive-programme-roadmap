import type { ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import type { TrackerData } from "../types/reporting";
import { parseDate } from "./dateUtils";

export type DateWindow = {
  start?: Date;
  end?: Date;
  label: string;
};

export type ExecutiveTone = "green" | "blue" | "amber" | "red" | "grey";

export type ExecutiveRoadmapPath = {
  outcome: ProgrammeItem;
  dependencies: ProgrammeItem[];
};

export type ExecutiveRoadmapModel = {
  reportDate: string;
  programmeStatus: string;
  originalDeliveryDate?: string;
  currentDeliveryDate?: string;
  forecastBasis: string;
  outcomes: ProgrammeItem[];
  paths: ExecutiveRoadmapPath[];
};

type ExecutivePathNode = {
  item: ProgrammeItem;
  depth: number;
};

export const executiveToneLabels: Record<ExecutiveTone, string> = {
  green: "GREEN",
  blue: "PLANNED",
  amber: "AMBER",
  red: "RED",
  grey: "NOT ASSESSED",
};

export function normaliseText(value?: string): string {
  return (value ?? "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

export function meaningfulText(value?: string): string | undefined {
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

export function bySoonest(a?: string, b?: string): number {
  return (parseDate(a)?.getTime() ?? Number.MAX_SAFE_INTEGER) - (parseDate(b)?.getTime() ?? Number.MAX_SAFE_INTEGER);
}

function weeklySummaryDate(summary: { meetingDate?: string; weekEnding?: string; lastUpdated?: string }): Date | undefined {
  return parseDate(summary.meetingDate) ?? parseDate(summary.weekEnding) ?? parseDate(summary.lastUpdated);
}

function latestWeeklySummary(tracker?: TrackerData) {
  return (tracker?.weeklySummaries ?? [])
    .slice()
    .sort((a, b) => (weeklySummaryDate(b)?.getTime() ?? 0) - (weeklySummaryDate(a)?.getTime() ?? 0))[0];
}

export function executiveMilestoneItems(schedule: ProgrammeSchedule): ProgrammeItem[] {
  const executive = schedule.items
    .filter((item) => item.executiveMilestone || normaliseText(item.milestoneLevel) === "executive milestone")
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
  if (executive.length) return executive;
  return schedule.items
    .filter((item) => item.boardReportable || item.roadmapMilestone || normaliseText(item.milestoneLevel).includes("executive"))
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate))
    .slice(0, 5);
}

function isDeliveredItem(item: ProgrammeItem): boolean {
  return item.status === "complete" || item.percentComplete === 100;
}

function isHistoricDeliveredItem(item: ProgrammeItem, windowStart?: Date): boolean {
  const finishDate = parseDate(item.finishDate);
  return Boolean(isDeliveredItem(item) && finishDate && windowStart && finishDate < windowStart);
}

export function programmeDeliveryOutcome(schedule: ProgrammeSchedule, outcomes: ProgrammeItem[]): ProgrammeItem | undefined {
  const candidates = outcomes.length ? outcomes : executiveMilestoneItems(schedule);
  return candidates.find((item) => /platform.*go live|go live/i.test(item.name))
    ?? candidates.slice().sort((a, b) => bySoonest(b.finishDate, a.finishDate))[0]
    ?? schedule.items.slice().sort((a, b) => bySoonest(b.finishDate, a.finishDate))[0];
}

function executiveDependencyScore(item: ProgrammeItem): number {
  const level = `${item.milestoneLevel ?? ""} ${item.dependencyLevel ?? ""}`.toLowerCase();
  let score = 0;
  if (item.executiveMilestone || normaliseText(item.milestoneLevel) === "executive milestone") score += 100;
  if (item.dependencyAnchor) score += 80;
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
  if (item.executiveMilestone || normaliseText(item.milestoneLevel) === "executive milestone") score += 100;
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

function collectPredecessorDependencies(outcome: ProgrammeItem, byUid: Map<string, ProgrammeItem>): ProgrammeItem[] {
  return collectPredecessorChainWithDepth(outcome, byUid)
    .filter((node) => node.depth === 1 || executivePathRelevance(node) >= 82)
    .sort((a, b) => executivePathRelevance(b) - executivePathRelevance(a) || bySoonest(a.item.finishDate, b.item.finishDate))
    .slice(0, 7)
    .sort((a, b) => bySoonest(a.item.finishDate, b.item.finishDate) || a.depth - b.depth)
    .map((node) => node.item);
}

export function executiveTone(item?: ProgrammeItem): ExecutiveTone {
  if (!item) return "grey";
  const rag = normaliseText(item.ragStatus);
  const confidence = normaliseText(item.dateConfidence);
  if (rag.includes("red")) return "red";
  if (rag.includes("amber")) return "amber";
  if (rag.includes("green")) return "green";
  if (item.status === "blocked") return "red";
  if (item.status === "complete") return "green";
  if (confidence.includes("high") || confidence.includes("confirmed") || confidence.includes("credible")) return "green";
  if (
    confidence.includes("medium") ||
    confidence.includes("low") ||
    confidence.includes("tbc") ||
    confidence.includes("unconfirmed") ||
    confidence.includes("assumption") ||
    item.externalDependency ||
    item.decisionRequired
  ) {
    return "amber";
  }
  if (item.finishDate) return "blue";
  return "grey";
}

export function buildExecutiveRoadmapModel(schedule: ProgrammeSchedule, tracker: TrackerData | undefined, dateWindow: DateWindow): ExecutiveRoadmapModel {
  const reportDate = new Date().toISOString();
  const windowStart = dateWindow.start ?? parseDate(reportDate);
  const allOutcomes = executiveMilestoneItems(schedule);
  const outcomes = allOutcomes.filter((item) => !isHistoricDeliveredItem(item, windowStart));
  const byUid = new Map(schedule.items.map((item) => [item.uid, item]));
  const deliveryOutcome = programmeDeliveryOutcome(schedule, allOutcomes);
  const weekly = latestWeeklySummary(tracker);
  return {
    reportDate,
    programmeStatus: meaningfulText(weekly?.overallRag) ?? "Not captured",
    originalDeliveryDate: deliveryOutcome?.baselineFinish,
    currentDeliveryDate: deliveryOutcome?.finishDate ?? schedule.finishDate,
    forecastBasis: deliveryOutcome?.name ?? "Programme finish",
    outcomes,
    paths: outcomes.map((outcome) => ({
      outcome,
      dependencies: collectPredecessorDependencies(outcome, byUid),
    })),
  };
}
