import * as XLSX from "xlsx";
import type {
  MeetingMinute,
  TrackerAction,
  TrackerChange,
  TrackerData,
  TrackerDecision,
  TrackerIssue,
  TrackerRisk,
  WeeklySummary,
} from "../types/reporting";

type Row = Record<string, string | number | boolean | Date | undefined>;

function clean(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  const text = String(value).trim();
  return text || undefined;
}

function normaliseHeader(value: unknown): string {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function excelDate(value: unknown): string | undefined {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(
      value.getFullYear(),
      value.getMonth(),
      value.getDate(),
      value.getHours(),
      value.getMinutes(),
      value.getSeconds(),
    )).toISOString();
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return undefined;
    return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.floor(parsed.S))).toISOString();
  }
  const text = clean(value);
  if (!text) return undefined;
  const ukDate = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:\s|$)/);
  if (ukDate) {
    const year = Number(ukDate[3].length === 2 ? `20${ukDate[3]}` : ukDate[3]);
    const month = Number(ukDate[2]);
    const day = Number(ukDate[1]);
    const parsed = new Date(Date.UTC(year, month - 1, day));
    if (parsed.getUTCFullYear() === year && parsed.getUTCMonth() === month - 1 && parsed.getUTCDate() === day) {
      return parsed.toISOString();
    }
  }
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? undefined : date.toISOString();
}

function yesNo(value: unknown): boolean {
  const text = clean(value)?.toLowerCase();
  return text === "yes" || text === "y" || text === "true" || text === "1" || text === "dashboard";
}

function makeRows(workbook: XLSX.WorkBook, sheetName: string): Row[] {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  const values = XLSX.utils.sheet_to_json<Array<string | number | boolean | Date | undefined>>(sheet, {
    header: 1,
    defval: undefined,
    raw: true,
  });
  const headerRowIndex = values.findIndex((row) => row.some((cell) => normaliseHeader(cell).endsWith("id") || normaliseHeader(cell) === "week ending"));
  if (headerRowIndex < 0) return [];
  const headers = values[headerRowIndex].map(normaliseHeader);
  return values.slice(headerRowIndex + 1).map((row) => {
    const record: Row = {};
    headers.forEach((header, index) => {
      if (!header) return;
      const incoming = row[index];
      const existing = record[header];
      const hasIncoming = incoming !== undefined && incoming !== null && String(incoming).trim() !== "";
      const hasExisting = existing !== undefined && existing !== null && String(existing).trim() !== "";
      if (!hasExisting || hasIncoming) record[header] = incoming;
    });
    return record;
  });
}

function rowValue(row: Row, ...names: string[]): string | undefined {
  for (const name of names) {
    const value = clean(row[normaliseHeader(name)]);
    if (value) return value;
  }
  return undefined;
}

function rowDate(row: Row, ...names: string[]): string | undefined {
  for (const name of names) {
    const value = excelDate(row[normaliseHeader(name)]);
    if (value) return value;
  }
  return undefined;
}

function rowFlag(row: Row, ...names: string[]): boolean {
  return names.some((name) => yesNo(row[normaliseHeader(name)]));
}

export async function parseMeetingTracker(file: File): Promise<TrackerData> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });

  const weeklySummaries: WeeklySummary[] = makeRows(workbook, "Weekly Summary")
    .map((row) => ({
      id: rowValue(row, "Summary ID") ?? "",
      weekEnding: rowDate(row, "Week ending"),
      meetingDate: rowDate(row, "Meeting date"),
      overallRag: rowValue(row, "Overall RAG"),
      ragRationale: rowValue(row, "RAG rationale"),
      topicsDiscussed: rowValue(row, "Topics discussed"),
      keyProgress: rowValue(row, "Key progress"),
      whatChanged: rowValue(row, "What changed this week"),
      keyRisksOrIssues: rowValue(row, "Key risks or issues"),
      decisionsMade: rowValue(row, "Decisions made"),
      decisionsNeeded: rowValue(row, "Decisions needed"),
      priorityActions: rowValue(row, "Priority actions"),
      executiveStatusSummary: rowValue(row, "Executive Status Summary", "Status Summary", "Executive Summary"),
      openingLine: rowValue(row, "Opening line for next meeting"),
      askSteerNeeded: rowValue(row, "Ask / Steer Needed", "CEO Ask"),
      mainBlocker: rowValue(row, "Main Blocker"),
      goLiveConfidence: rowValue(row, "Go Live Confidence"),
      ragMovement: rowValue(row, "RAG Movement"),
      steerRequired: rowValue(row, "Steer Required", "Decision Required From CEO"),
      updateType: rowValue(row, "Update type"),
      lastUpdated: rowDate(row, "Last updated"),
    }))
    .filter((item) => item.id || item.weekEnding);

  const risks: TrackerRisk[] = makeRows(workbook, "Risks")
    .map((row) => ({
      id: rowValue(row, "Risk ID") ?? "",
      dateRaised: rowDate(row, "Date raised"),
      dashboardFlag: rowFlag(row, "Dashboard Tag", "Dashboard Flag"),
      lastDiscussedDate: rowDate(row, "Last discussed date"),
      stream: rowValue(row, "Workstream"),
      title: rowValue(row, "Risk title") ?? "",
      statement: rowValue(row, "Risk statement"),
      cause: rowValue(row, "Cause"),
      impact: rowValue(row, "Impact"),
      likelihood: rowValue(row, "Likelihood"),
      impactRating: rowValue(row, "Impact rating"),
      rag: rowValue(row, "RAG"),
      mitigation: rowValue(row, "Mitigation"),
      owner: rowValue(row, "Owner"),
      targetDate: rowDate(row, "Target date"),
      status: rowValue(row, "Status"),
      updateType: rowValue(row, "Update type"),
      latestUpdate: rowValue(row, "Latest update"),
    }))
    .filter((item) => item.id || item.title);

  const issues: TrackerIssue[] = makeRows(workbook, "Issues")
    .map((row) => ({
      id: rowValue(row, "Issue ID") ?? "",
      dateRaised: rowDate(row, "Date raised"),
      dashboardFlag: rowFlag(row, "Dashboard Tag", "Dashboard Flag"),
      lastDiscussedDate: rowDate(row, "Last discussed date"),
      stream: rowValue(row, "Workstream"),
      title: rowValue(row, "Issue title") ?? "",
      statement: rowValue(row, "Issue statement"),
      impact: rowValue(row, "Impact"),
      priority: rowValue(row, "Priority"),
      rag: rowValue(row, "RAG"),
      requiredAction: rowValue(row, "Required action"),
      requiredDecision: rowValue(row, "Required decision"),
      owner: rowValue(row, "Owner"),
      targetDate: rowDate(row, "Target date"),
      status: rowValue(row, "Status"),
      updateType: rowValue(row, "Update type"),
      latestUpdate: rowValue(row, "Latest update"),
    }))
    .filter((item) => item.id || item.title);

  const actions: TrackerAction[] = makeRows(workbook, "Actions")
    .map((row) => ({
      id: rowValue(row, "Action ID") ?? "",
      meetingDate: rowDate(row, "Meeting date", "Meeting/Log date", "Log date", "Logged date", "Date logged", "Action logged date"),
      dashboardFlag: rowFlag(row, "Dashboard Tag", "Dashboard Flag"),
      stream: rowValue(row, "Workstream"),
      title: rowValue(row, "Action title", "Action", "Task", "Task title") ?? "",
      description: rowValue(row, "Task description", "Action description", "Description"),
      status: rowValue(row, "Status"),
      owner: rowValue(row, "Owner", "Assigned to", "Action owner"),
      priority: rowValue(row, "Priority"),
      dueDate: rowDate(row, "Due date", "Due Date", "Action due date", "Date due", "Target date", "Target completion date"),
      completionDate: rowDate(row, "Completion date", "Completed date", "Date completed", "Closed date"),
      updateType: rowValue(row, "Update type"),
      latestUpdate: rowValue(row, "Latest update"),
    }))
    .filter((item) => item.id || item.title);

  const decisions: TrackerDecision[] = makeRows(workbook, "Decisions")
    .map((row) => ({
      id: rowValue(row, "Decision ID") ?? "",
      decisionDate: rowDate(row, "Decision date"),
      decisionRequiredBy: rowDate(row, "Decision required by"),
      decisionRequiredByLabel: rowValue(row, "Decision required by"),
      dashboardFlag: rowFlag(row, "Dashboard Tag", "Dashboard Flag"),
      decisionType: rowValue(row, "Decision Type"),
      lastDiscussedDate: rowDate(row, "Last discussed date"),
      stream: rowValue(row, "Workstream"),
      title: rowValue(row, "Decision title") ?? "",
      statement: rowValue(row, "Decision statement"),
      decisionMaker: rowValue(row, "Decision maker"),
      owner: rowValue(row, "Owner"),
      status: rowValue(row, "Status"),
      updateType: rowValue(row, "Update type"),
      latestUpdate: rowValue(row, "Latest update"),
    }))
    .filter((item) => item.id || item.title);

  const changes: TrackerChange[] = makeRows(workbook, "Changes")
    .map((row) => ({
      id: rowValue(row, "Change ID") ?? "",
      dateRaised: rowDate(row, "Date raised"),
      lastDiscussedDate: rowDate(row, "Last discussed date"),
      dashboardFlag: rowFlag(row, "Dashboard Tag", "Dashboard Flag"),
      changeType: rowValue(row, "Change Type"),
      stream: rowValue(row, "Workstream"),
      title: rowValue(row, "Change title") ?? "",
      description: rowValue(row, "Change description"),
      reason: rowValue(row, "Reason for change"),
      impactOnTime: rowValue(row, "Impact on time"),
      impactOnScope: rowValue(row, "Impact on scope"),
      impactOnCost: rowValue(row, "Impact on cost"),
      impactOnQualityOrBenefits: rowValue(row, "Impact on quality or Benefits"),
      decisionRequired: rowValue(row, "Decision required"),
      decisionMaker: rowValue(row, "Decision maker"),
      owner: rowValue(row, "Owner"),
      status: rowValue(row, "Status"),
      updateType: rowValue(row, "Update type"),
      latestUpdate: rowValue(row, "Latest update"),
    }))
    .filter((item) => item.id || item.title);

  const minutes: MeetingMinute[] = makeRows(workbook, "Meeting Minutes")
    .map((row) => ({
      id: rowValue(row, "Minute ID") ?? "",
      meetingDate: rowDate(row, "Meeting date"),
      meetingName: rowValue(row, "Meeting name"),
      stream: rowValue(row, "Workstream"),
      topic: rowValue(row, "Topic"),
      summary: rowValue(row, "Discussion summary"),
      outcomeType: rowValue(row, "Outcome type"),
      updateType: rowValue(row, "Update type"),
      latestUpdate: rowValue(row, "Latest update"),
    }))
    .filter((item) => item.id || item.summary);

  return {
    sourceFileName: file.name,
    importedAt: new Date().toISOString(),
    weeklySummaries,
    risks,
    issues,
    actions,
    decisions,
    changes,
    minutes,
  };
}
