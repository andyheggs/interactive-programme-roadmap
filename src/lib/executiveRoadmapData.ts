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

export type ExecutiveToneAssessment = {
  tone: ExecutiveTone;
  summary: string;
  reasons: string[];
  evidence: Array<{ label: string; value: string }>;
};

export const executiveToneLabels: Record<ExecutiveTone, string> = {
  green: "GREEN",
  blue: "PLANNED",
  amber: "AMBER",
  red: "RED",
  grey: "NOT ASSESSED",
};

const uncertainDateTerms = [
  "medium",
  "low",
  "tbc",
  "unconfirmed",
  "assumption",
  "not yet confirmed",
  "unknown",
  "subject to",
  "dependent on",
  "supplier implementation",
  "supplier plan",
];

const confirmedDateTerms = [
  "actual",
  "confirmed",
  "credible",
  "high",
  "statutory",
  "fixed",
  "agreed",
];

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

function includesAny(value: string, terms: string[]): boolean {
  return terms.some((term) => value.includes(term));
}

function summariseAssessment(tone: ExecutiveTone, reasons: string[]): string {
  const firstReason = reasons[0]?.replace(/\.$/, "");
  if (tone === "amber") return firstReason ? `Amber because ${firstReason.toLowerCase()}.` : "Amber because the date or delivery position is uncertain.";
  if (tone === "red") return firstReason ? `Red because ${firstReason.toLowerCase()}.` : "Red because the item is blocked, overdue or reported Red.";
  if (tone === "green") return firstReason ? `Green because ${firstReason.toLowerCase()}.` : "Green because the item is complete, confirmed or reported Green.";
  if (tone === "blue") return firstReason ? `Planned because ${firstReason.toLowerCase()}.` : "Planned because it has a dated forecast but no RAG concern.";
  return firstReason ? `Not assessed because ${firstReason.toLowerCase()}.` : "Not assessed because no RAG or date confidence is captured.";
}

export function executiveToneAssessment(item?: ProgrammeItem): ExecutiveToneAssessment {
  if (!item) {
    return {
      tone: "grey",
      summary: "Not assessed because no project plan item is selected.",
      reasons: ["No project plan item is selected."],
      evidence: [],
    };
  }

  const rag = normaliseText(item.ragStatus);
  const confidence = normaliseText(item.dateConfidence);
  const confirmedDate = includesAny(confidence, confirmedDateTerms);
  const uncertainDate = Boolean(includesAny(confidence, uncertainDateTerms) || (!confirmedDate && confidence.includes("target")));
  const evidence = [
    meaningfulText(item.ragStatus) ? { label: "Plan RAG", value: meaningfulText(item.ragStatus)! } : undefined,
    meaningfulText(item.dateConfidence) ? { label: "Date confidence", value: meaningfulText(item.dateConfidence)! } : undefined,
    item.externalDependency ? { label: "External dependency", value: "Yes" } : undefined,
    item.decisionRequired ? { label: "Decision required", value: "Yes" } : undefined,
  ].filter((entry): entry is { label: string; value: string } => Boolean(entry));

  let tone: ExecutiveTone = "grey";
  const reasons: string[] = [];

  if (rag.includes("red")) {
    tone = "red";
    reasons.push("Source RAG is Red");
  } else if (item.status === "blocked") {
    tone = "red";
    reasons.push("Project status is blocked");
  } else if (rag.includes("amber")) {
    tone = "amber";
    reasons.push("Source RAG is Amber");
  } else if (uncertainDate) {
    tone = "amber";
    reasons.push(item.dateConfidence ? `Date confidence is "${item.dateConfidence}"` : "The forecast date is not confirmed");
  } else if (item.externalDependency && !confirmedDate && !rag.includes("green")) {
    tone = "amber";
    reasons.push("Delivery depends on an external item and no confirmed date confidence is captured");
  } else if (item.status === "complete") {
    tone = "green";
    reasons.push("Project status is complete");
  } else if (rag.includes("green")) {
    tone = "green";
    reasons.push("Source RAG is Green");
  } else if (confirmedDate) {
    tone = "green";
    reasons.push(item.dateConfidence ? `Date confidence is "${item.dateConfidence}"` : "Date confidence is confirmed");
  } else if (item.finishDate) {
    tone = "blue";
    reasons.push("A forecast finish date is captured and no RAG concern is flagged");
  } else {
    reasons.push("No RAG status or date confidence is captured");
  }

  if (item.decisionRequired) reasons.push("Decision gate is flagged in the Project plan");
  if (item.externalDependency && !reasons.some((reason) => reason.toLowerCase().includes("external"))) {
    reasons.push("External dependency is flagged in the Project plan");
  }

  return {
    tone,
    summary: summariseAssessment(tone, reasons),
    reasons,
    evidence,
  };
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

function collectPredecessorDependencies(outcome: ProgrammeItem, byUid: Map<string, ProgrammeItem>): ProgrammeItem[] {
  return collectPredecessorChainWithDepth(outcome, byUid)
    .filter((node) => visibleExecutivePathItem(node.item))
    .sort((a, b) => executivePathRelevance(b) - executivePathRelevance(a) || bySoonest(a.item.finishDate, b.item.finishDate))
    .slice(0, 7)
    .sort((a, b) => bySoonest(a.item.finishDate, b.item.finishDate) || a.depth - b.depth)
    .map((node) => node.item);
}

export function executiveTone(item?: ProgrammeItem): ExecutiveTone {
  return executiveToneAssessment(item).tone;
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
