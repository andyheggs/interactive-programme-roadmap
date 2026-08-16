import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  AlertTriangle,
  CalendarDays,
  CheckCircle2,
  ChevronRight,
  ClipboardCheck,
  Clock,
  Diamond,
  Download,
  FileUp,
  FileSpreadsheet,
  Filter,
  GitBranch,
  GripVertical,
  Layers,
  LayoutDashboard,
  Moon,
  Milestone,
  Search,
  ShieldCheck,
  SlidersHorizontal,
  Sun,
  Target,
  Users,
  X,
} from "lucide-react";
import { exportProgrammePdf } from "./lib/exportProgrammePdf";
import { exportExecutiveRoadmapPdf } from "./lib/exportExecutiveRoadmapPdf";
import {
  exportExecutiveRoadmapHtml,
  exportExecutiveRoadmapImage,
  exportExecutiveRoadmapPosterPdf,
} from "./lib/exportExecutiveRoadmapVisuals";
import { exportGanttPdf, type GanttPdfSection } from "./lib/exportGanttPdf";
import { buildExecutiveRoadmapModel, executiveToneAssessment, executiveToneLabel } from "./lib/executiveRoadmapData";
import { exportTeamActionsPdf as exportTeamActionsA4Pdf } from "./lib/exportTeamActionsPdf";
import { exportWeeklyStatusPdf as exportWeeklyStatusA4Pdf } from "./lib/exportWeeklyStatusPdf";
import { parseMicrosoftProjectXml } from "./lib/parseMicrosoftProjectXml";
import { parseMeetingTracker } from "./lib/parseMeetingTracker";
import { clamp, durationLabel, formatDate, parseDate, uniqueSorted } from "./lib/dateUtils";
import type { ProgrammeFilters, ProgrammeItem, ProgrammeSchedule, ProgrammeView } from "./types/programme";
import type { TrackerAction, TrackerChange, TrackerData, TrackerDecision, TrackerIssue, TrackerRisk, WeeklyStatusCuration, WeeklyStatusSectionKey, WeeklySummary } from "./types/reporting";
import "./styles.css";

const initialFilters: ProgrammeFilters = {
  stream: "all",
  roadmapView: "all",
  milestoneType: "all",
  approvalBody: "all",
  version: "all",
  visibility: "all",
  status: "all",
  criticalOnly: false,
  roadmapOnly: false,
  showSummaryTasks: true,
  delayedOnly: false,
  datePreset: "all",
  dateStart: "",
  dateEnd: "",
  search: "",
};

const viewLabels: Record<ProgrammeView, string> = {
  roadmap: "Programme Roadmap",
  schedule: "Integrated Schedule",
  milestones: "Milestones",
  governance: "Governance",
  delivery: "Delivery",
  release: "Version / Release",
};

type AppPage =
  | "workspace"
  | "home"
  | "ceo"
  | "weekly-status"
  | "team-actions"
  | "board"
  | "reporting-roadmap"
  | "risks"
  | "actions"
  | "gantt"
  | "release-roadmap"
  | "version-scope"
  | "release-readiness"
  | "dependencies"
  | "workstreams"
  | "partner"
  | "downloads";

const appPages: Array<{ key: AppPage; label: string; group: "core" | "reporting" | "future" }> = [
  { key: "workspace", label: "Roadmap Workspace", group: "core" },
  { key: "home", label: "Home Dashboard", group: "reporting" },
  { key: "ceo", label: "Executive View", group: "reporting" },
  { key: "weekly-status", label: "Weekly Executive Status", group: "reporting" },
  { key: "team-actions", label: "Team Action Tracker", group: "reporting" },
  { key: "board", label: "Board Report", group: "reporting" },
  { key: "reporting-roadmap", label: "Reporting Roadmap", group: "reporting" },
  { key: "gantt", label: "Gantt View", group: "reporting" },
  { key: "risks", label: "Risks & Issues", group: "reporting" },
  { key: "actions", label: "Actions & Decisions", group: "reporting" },
  { key: "dependencies", label: "Dependencies", group: "reporting" },
  { key: "workstreams", label: "Workstreams", group: "reporting" },
  { key: "partner", label: "Partner View", group: "reporting" },
  { key: "downloads", label: "Downloads", group: "reporting" },
  { key: "release-roadmap", label: "Release Roadmap", group: "future" },
  { key: "version-scope", label: "Version Scope", group: "future" },
  { key: "release-readiness", label: "Release Readiness", group: "future" },
];

function makeSampleSchedule(): ProgrammeSchedule {
  const items: ProgrammeItem[] = [
    {
      uid: "1",
      id: "1",
      name: "Upload a Microsoft Project XML file to generate the roadmap",
      outlineLevel: 1,
      itemType: "summary",
      isSummary: true,
      isMilestone: false,
      isCritical: false,
      isActive: true,
      startDate: "2026-01-09T08:00:00",
      finishDate: "2027-03-11T17:00:00",
      percentComplete: 0,
      predecessors: [],
      successors: [],
      stream: "Programme Controls",
      visibility: "Internal",
      status: "not-started",
    },
  ];
  return {
    title: "Interactive Programme Roadmap",
    items,
    resources: [],
    importedAt: new Date().toISOString(),
  };
}

function applyView(items: ProgrammeItem[], view: ProgrammeView): ProgrammeItem[] {
  if (view === "roadmap") return items.filter((item) => item.isSummary || item.roadmapMilestone || item.outlineLevel <= 2);
  if (view === "milestones") return items.filter((item) => item.isMilestone || item.roadmapMilestone);
  if (view === "governance") return items.filter((item) => item.approvalBody || item.roadmapView === "Governance" || item.milestoneType === "Approval");
  if (view === "delivery") return items.filter((item) => item.roadmapView === "Delivery" || Boolean(item.stream));
  if (view === "release") return items.filter((item) => Boolean(item.version));
  return items;
}

type DateWindow = {
  start?: Date;
  end?: Date;
  label: string;
};

function addDays(date: Date, days: number): Date {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function inputDate(value?: string): string {
  const date = parseDate(value);
  return date ? date.toISOString().slice(0, 10) : "";
}

function resolveDateWindow(filters: ProgrammeFilters, schedule: ProgrammeSchedule): DateWindow {
  const statusDate = parseDate(schedule.statusDate);
  const currentDate = parseDate(schedule.currentDate) ?? statusDate;
  if (filters.datePreset === "status-forward" && statusDate) return { start: statusDate, label: "Status date forward" };
  if (filters.datePreset === "current-forward" && currentDate) return { start: currentDate, label: "Current date forward" };
  if (filters.datePreset === "next-30" && statusDate) return { start: statusDate, end: addDays(statusDate, 30), label: "Next 30 days" };
  if (filters.datePreset === "next-60" && statusDate) return { start: statusDate, end: addDays(statusDate, 60), label: "Next 60 days" };
  if (filters.datePreset === "next-90" && statusDate) return { start: statusDate, end: addDays(statusDate, 90), label: "Next 90 days" };
  if (filters.datePreset === "custom") {
    return {
      start: filters.dateStart ? parseDate(`${filters.dateStart}T00:00:00`) : undefined,
      end: filters.dateEnd ? parseDate(`${filters.dateEnd}T23:59:59`) : undefined,
      label: "Custom date window",
    };
  }
  return { label: "Full programme" };
}

function overlapsDateWindow(item: ProgrammeItem, window: DateWindow): boolean {
  if (!window.start && !window.end) return true;
  const itemStart = parseDate(item.startDate) ?? parseDate(item.finishDate);
  const itemFinish = parseDate(item.finishDate) ?? parseDate(item.startDate);
  if (!itemStart || !itemFinish) return false;
  if (window.start && itemFinish < window.start) return false;
  if (window.end && itemStart > window.end) return false;
  return true;
}

function applyFilters(items: ProgrammeItem[], filters: ProgrammeFilters, schedule: ProgrammeSchedule): ProgrammeItem[] {
  const search = filters.search.trim().toLowerCase();
  const dateWindow = resolveDateWindow(filters, schedule);
  return items.filter((item) => {
    if (filters.stream !== "all" && item.stream !== filters.stream) return false;
    if (filters.roadmapView !== "all" && item.roadmapView !== filters.roadmapView) return false;
    if (filters.milestoneType !== "all" && item.milestoneType !== filters.milestoneType) return false;
    if (filters.approvalBody !== "all" && item.approvalBody !== filters.approvalBody) return false;
    if (filters.version !== "all" && item.version !== filters.version) return false;
    if (filters.visibility !== "all" && item.visibility !== filters.visibility) return false;
    if (filters.status !== "all" && item.status !== filters.status) return false;
    if (filters.criticalOnly && !item.isCritical) return false;
    if (filters.roadmapOnly && !item.roadmapMilestone) return false;
    if (!filters.showSummaryTasks && item.isSummary) return false;
    if (filters.delayedOnly && !(item.delayDays && item.delayDays > 0)) return false;
    if (!overlapsDateWindow(item, dateWindow)) return false;
    if (search && !`${item.name} ${item.stream ?? ""} ${item.approvalBody ?? ""}`.toLowerCase().includes(search)) return false;
    return true;
  });
}

function timelineBounds(items: ProgrammeItem[], schedule: ProgrammeSchedule, dateWindow?: DateWindow) {
  const dates = [
    dateWindow?.start ?? parseDate(schedule.startDate),
    dateWindow?.end ?? parseDate(schedule.finishDate),
    ...items.flatMap((item) => [parseDate(item.startDate), parseDate(item.finishDate), parseDate(item.baselineFinish)]),
  ].filter((date): date is Date => Boolean(date));
  const rawMin = new Date(Math.min(...dates.map((date) => date.getTime())));
  const rawMax = new Date(Math.max(...dates.map((date) => date.getTime())));
  const min = dateWindow?.start ?? rawMin;
  const max = dateWindow?.end ?? rawMax;
  return { min, max, span: Math.max(1, max.getTime() - min.getTime()) };
}

function positionFor(dateValue: string | undefined, bounds: ReturnType<typeof timelineBounds>) {
  const date = parseDate(dateValue);
  if (!date) return 0;
  return clamp(((date.getTime() - bounds.min.getTime()) / bounds.span) * 100, 0, 100);
}

function SummaryBar({ schedule, tracker }: { schedule: ProgrammeSchedule; tracker?: TrackerData }) {
  const totalMilestones = schedule.items.filter((item) => item.isMilestone).length;
  const roadmapMilestones = schedule.items.filter((item) => item.roadmapMilestone).length;
  const delayed = schedule.items.filter((item) => item.delayDays && item.delayDays > 0).length;
  const latestWeekly = latestWeeklySummary(tracker);
  const reportDate = latestWeekly?.meetingDate ?? latestWeekly?.weekEnding ?? new Date().toISOString();
  const cards = [
    ["Programme start", formatDate(schedule.startDate)],
    ["Programme finish", formatDate(schedule.finishDate)],
    ["Report date", formatDate(reportDate)],
    ["Schedule status", formatDate(schedule.statusDate)],
    ["Total tasks", schedule.items.length.toString()],
    ["Milestones", totalMilestones.toString()],
    ["Roadmap", roadmapMilestones.toString()],
    ["Delayed", delayed.toString()],
  ];
  return (
    <section className="summary-bar">
      {cards.map(([label, value]) => (
        <div className="metric" key={label}>
          <span>{label}</span>
          <strong>{value}</strong>
        </div>
      ))}
    </section>
  );
}

function SelectFilter({
  label,
  value,
  values,
  onChange,
}: {
  label: string;
  value: string;
  values: string[];
  onChange: (value: string) => void;
}) {
  return (
    <label className="field">
      <span>{label}</span>
      <select value={value} onChange={(event) => onChange(event.target.value)}>
        <option value="all">All</option>
        {values.map((option) => (
          <option value={option} key={option}>
            {option}
          </option>
        ))}
      </select>
    </label>
  );
}

function FilterPanel({
  schedule,
  filters,
  setFilters,
}: {
  schedule: ProgrammeSchedule;
  filters: ProgrammeFilters;
  setFilters: React.Dispatch<React.SetStateAction<ProgrammeFilters>>;
}) {
  const options = useMemo(
    () => ({
      streams: uniqueSorted(schedule.items.map((item) => item.stream)),
      roadmapViews: uniqueSorted(schedule.items.map((item) => item.roadmapView)),
      milestoneTypes: uniqueSorted(schedule.items.map((item) => item.milestoneType)),
      approvalBodies: uniqueSorted(schedule.items.map((item) => item.approvalBody)),
      versions: uniqueSorted(schedule.items.map((item) => item.version)),
      visibility: uniqueSorted(schedule.items.map((item) => item.visibility)),
      statuses: uniqueSorted(schedule.items.map((item) => item.status)),
    }),
    [schedule],
  );
  const update = <K extends keyof ProgrammeFilters>(key: K, value: ProgrammeFilters[K]) =>
    setFilters((current) => ({ ...current, [key]: value }));

  return (
    <aside className="side-panel">
      <div className="panel-title">
        <Filter size={18} />
        <strong>Filters</strong>
      </div>
      <label className="field search-field">
        <span>Search</span>
        <Search size={16} />
        <input value={filters.search} onChange={(event) => update("search", event.target.value)} placeholder="Task, stream, approval..." />
      </label>
      <SelectFilter label="Stream" value={filters.stream} values={options.streams} onChange={(value) => update("stream", value)} />
      <SelectFilter label="Roadmap view" value={filters.roadmapView} values={options.roadmapViews} onChange={(value) => update("roadmapView", value)} />
      <SelectFilter label="Milestone type" value={filters.milestoneType} values={options.milestoneTypes} onChange={(value) => update("milestoneType", value)} />
      <SelectFilter label="Approval body" value={filters.approvalBody} values={options.approvalBodies} onChange={(value) => update("approvalBody", value)} />
      <SelectFilter label="Version" value={filters.version} values={options.versions} onChange={(value) => update("version", value)} />
      <SelectFilter label="Visibility" value={filters.visibility} values={options.visibility} onChange={(value) => update("visibility", value)} />
      <SelectFilter label="Status" value={filters.status} values={options.statuses} onChange={(value) => update("status", value)} />
      <label className="field">
        <span>Date window</span>
        <select value={filters.datePreset} onChange={(event) => update("datePreset", event.target.value as ProgrammeFilters["datePreset"])}>
          <option value="all">Full programme</option>
          <option value="status-forward">Status date forward</option>
          <option value="current-forward">Current date forward</option>
          <option value="next-30">Next 30 days</option>
          <option value="next-60">Next 60 days</option>
          <option value="next-90">Next 90 days</option>
          <option value="custom">Custom</option>
        </select>
      </label>
      {filters.datePreset === "custom" ? (
        <div className="date-range">
          <label className="field">
            <span>From</span>
            <input type="date" value={filters.dateStart} onChange={(event) => update("dateStart", event.target.value)} />
          </label>
          <label className="field">
            <span>To</span>
            <input type="date" value={filters.dateEnd} onChange={(event) => update("dateEnd", event.target.value)} />
          </label>
        </div>
      ) : null}
      <label className="check"><input type="checkbox" checked={filters.criticalOnly} onChange={(event) => update("criticalOnly", event.target.checked)} /> Critical only</label>
      <label className="check"><input type="checkbox" checked={filters.roadmapOnly} onChange={(event) => update("roadmapOnly", event.target.checked)} /> Roadmap milestones only</label>
      <label className="check"><input type="checkbox" checked={filters.showSummaryTasks} onChange={(event) => update("showSummaryTasks", event.target.checked)} /> Show summary tasks</label>
      <label className="check"><input type="checkbox" checked={filters.delayedOnly} onChange={(event) => update("delayedOnly", event.target.checked)} /> Delayed only</label>
    </aside>
  );
}

function Timeline({
  schedule,
  items,
  selected,
  onSelect,
  dateWindow,
}: {
  schedule: ProgrammeSchedule;
  items: ProgrammeItem[];
  selected?: ProgrammeItem;
  onSelect: (item: ProgrammeItem) => void;
  dateWindow: DateWindow;
}) {
  const bounds = useMemo(() => timelineBounds(items.length ? items : schedule.items, schedule, dateWindow), [dateWindow, items, schedule]);
  const visibleItems = items.slice(0, 180);
  const statusMarker = positionFor(schedule.statusDate, bounds);
  return (
    <section className="timeline-shell">
      <div className="timeline-toolbar">
        <div>
          <strong>{visibleItems.length}</strong> items shown
          {items.length > visibleItems.length ? <span> from {items.length} filtered results</span> : null}
          <span className="window-label"> {dateWindow.label}</span>
        </div>
        <div className="legend">
          <span><i className="key summary" />Summary</span>
          <span><i className="key task" />Task</span>
          <span><Diamond size={13} />Milestone</span>
          <span><Diamond className="legend-roadmap" size={13} />Roadmap milestone</span>
          <span><i className="key complete" />Complete</span>
          <span><i className="key progress" />In progress</span>
          <span><i className="key risk" />At risk</span>
          <span><i className="key late" />Late / delayed</span>
          <span><i className="key critical" />Critical</span>
          <span><i className="key baseline" />Baseline</span>
          <span><i className="key status" />Status date</span>
        </div>
      </div>
      <div className="timeline-scale">
        <span>{formatDate(bounds.min.toISOString())}</span>
        <span>Status date</span>
        <span>{formatDate(bounds.max.toISOString())}</span>
      </div>
      <div className="timeline" style={{ "--status-left": `${statusMarker}%` } as React.CSSProperties}>
        {visibleItems.map((item) => {
          const start = positionFor(item.startDate, bounds);
          const finish = positionFor(item.finishDate, bounds);
          const width = Math.max(item.isMilestone ? 0.8 : 1.2, finish - start);
          const baseline = positionFor(item.baselineFinish, bounds);
          return (
            <button
              type="button"
              className={`timeline-row ${selected?.uid === item.uid ? "selected" : ""}`}
              key={item.uid}
              onClick={() => onSelect(item)}
            >
              <span className="row-label" style={{ paddingLeft: `${Math.min(item.outlineLevel - 1, 5) * 14}px` }}>
                {item.childUids?.length ? <ChevronRight size={14} /> : <span className="indent-spacer" />}
                <span className="row-name">{item.name}</span>
              </span>
              <span className="row-track">
                {item.baselineFinish ? <span className="baseline-marker" style={{ left: `${baseline}%` }} /> : null}
                {item.isMilestone ? (
                  <span className={`milestone ${item.roadmapMilestone ? "roadmap" : ""} ${item.isCritical ? "critical" : ""}`} style={{ left: `${finish}%` }} />
                ) : (
                  <span
                    className={`bar ${item.itemType} ${item.status} ${item.isCritical ? "critical" : ""}`}
                    style={{ left: `${start}%`, width: `${width}%` }}
                  >
                    <span style={{ width: `${item.percentComplete ?? 0}%` }} />
                  </span>
                )}
                {item.delayDays && item.delayDays > 0 ? <em className="delay-badge" style={{ left: `${Math.min(96, finish + 1)}%` }}>+{item.delayDays}d</em> : null}
              </span>
            </button>
          );
        })}
      </div>
    </section>
  );
}

function InsightsPanel({ schedule }: { schedule: ProgrammeSchedule }) {
  const status = parseDate(schedule.statusDate) ?? new Date();
  const next14 = schedule.items
    .filter((item) => {
      const start = parseDate(item.startDate);
      return start && start >= status && start.getTime() - status.getTime() <= 14 * 86_400_000;
    })
    .slice(0, 5);
  const delayed = schedule.items.filter((item) => item.delayDays && item.delayDays > 0).sort((a, b) => (b.delayDays ?? 0) - (a.delayDays ?? 0)).slice(0, 5);
  const approvals = schedule.items.filter((item) => item.approvalBody && (item.isMilestone || item.roadmapMilestone)).slice(0, 5);
  const critical = schedule.items.filter((item) => item.isCritical && item.status !== "complete").slice(0, 5);
  const groups = [
    ["Next 14 days", next14],
    ["Late / slipping", delayed],
    ["Approval decisions", approvals],
    ["Critical open items", critical],
  ] as const;
  return (
    <section className="insights">
      {groups.map(([title, group]) => (
        <div className="insight" key={title}>
          <h3>{title}</h3>
          {group.length ? group.map((item) => <p key={item.uid}><strong>{formatDate(item.finishDate)}</strong> {item.name}</p>) : <p>No items found.</p>}
        </div>
      ))}
    </section>
  );
}

function isOpenStatus(status?: string): boolean {
  const value = status?.toLowerCase() ?? "";
  return !["complete", "completed", "closed", "done"].includes(value);
}

function isRedOrAmber(value?: string): boolean {
  const label = value?.toLowerCase() ?? "";
  return label.includes("red") || label.includes("amber") || label.includes("high");
}

function dateWithin(value: string | undefined, window: DateWindow): boolean {
  if (!window.start && !window.end) return true;
  const date = parseDate(value);
  if (!date) return false;
  if (window.start && date < window.start) return false;
  if (window.end && date > window.end) return false;
  return true;
}

function bySoonest(a?: string, b?: string): number {
  return (parseDate(a)?.getTime() ?? Number.MAX_SAFE_INTEGER) - (parseDate(b)?.getTime() ?? Number.MAX_SAFE_INTEGER);
}

function toneClass(value?: string): "red" | "amber" | "green" | "neutral" {
  const text = value?.toLowerCase() ?? "";
  if (text.includes("red") || text.includes("high")) return "red";
  if (text.includes("amber") || text.includes("medium")) return "amber";
  if (text.includes("green") || text.includes("low")) return "green";
  return "neutral";
}

function sortFlaggedFirst<T extends { dashboardFlag?: boolean }>(items: T[]): T[] {
  return items.slice().sort((a, b) => Number(Boolean(b.dashboardFlag)) - Number(Boolean(a.dashboardFlag)));
}

function meaningfulText(value?: string): string | undefined {
  const text = value?.trim();
  if (!text) return undefined;
  const normalised = normaliseText(text);
  if (!normalised) return undefined;
  if (
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

function formatDateOrText(value?: string, fallback = "Not set"): string {
  if (!value) return fallback;
  return parseDate(value) ? formatDate(value) : value;
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

function itemImportance(item: ProgrammeItem): number {
  const level = item.milestoneLevel?.toLowerCase() ?? "";
  if (item.executiveMilestone || level.includes("executive")) return 5;
  if (item.boardReportable || level.includes("board")) return 4;
  if (item.roadmapMilestone) return 3;
  if (item.governanceGate || item.decisionRequired) return 2;
  return 1;
}

function latestWeeklySummary(tracker?: TrackerData) {
  return sortedWeeklySummaries(tracker)[0];
}

function weeklySummaryDate(summary: { meetingDate?: string; weekEnding?: string; lastUpdated?: string }): Date | undefined {
  return parseDate(summary.meetingDate) ?? parseDate(summary.weekEnding) ?? parseDate(summary.lastUpdated);
}

function sortedWeeklySummaries(tracker?: TrackerData) {
  return (tracker?.weeklySummaries ?? [])
    .slice()
    .sort((a, b) => {
      const aDate = weeklySummaryDate(a);
      const bDate = weeklySummaryDate(b);
      return (bDate?.getTime() ?? 0) - (aDate?.getTime() ?? 0);
    });
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

function isCompleteStatus(status?: string): boolean {
  return ["complete", "completed", "closed", "done", "resolved", "implemented"].includes(normaliseText(status));
}

function isBlockedStatus(status?: string): boolean {
  const value = normaliseText(status);
  return value.includes("block") || value.includes("hold");
}

function actionStatusGroup(item: TeamWorkItem): "open" | "due-soon" | "overdue" | "blocked" | "completed" {
  if (isCompleteStatus(item.status)) return "completed";
  if (isBlockedStatus(item.status)) return "blocked";
  const due = parseDate(item.dueDate);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (due && due < today) return "overdue";
  if (due && due <= addDays(today, 14)) return "due-soon";
  return "open";
}

function statusLabel(item: TeamWorkItem): string {
  const group = actionStatusGroup(item);
  if (group === "due-soon") return item.status ? `${item.status} / due soon` : "Due soon";
  if (group === "overdue") return item.status ? `${item.status} / overdue` : "Overdue";
  return item.status ?? "Not set";
}

function teamStatusMatches(group: ReturnType<typeof actionStatusGroup>, filters: TeamStatusFilter[]): boolean {
  if (!filters.length) return true;
  return filters.includes(group);
}

function combineTeamWorkItems(schedule: ProgrammeSchedule, tracker?: TrackerData): TeamWorkItem[] {
  const trackerItems: TeamWorkItem[] = (tracker?.actions ?? []).map((action) => ({
    id: `tracker-${action.id}`,
    source: "Meeting action",
    title: action.title,
    owner: action.owner,
    stream: action.stream,
    status: action.status,
    priority: action.priority,
    loggedDate: action.meetingDate,
    dueDate: action.dueDate,
    completionDate: action.completionDate,
    meetingDate: action.meetingDate,
    description: action.description,
    latestUpdate: action.latestUpdate,
    links: [action.id, action.updateType].filter(Boolean).join(" · "),
    dashboardFlag: action.dashboardFlag,
  }));
  const projectItems: TeamWorkItem[] = schedule.items
    .filter((item) => !item.isSummary && item.isActive && item.resourceNames?.length)
    .map((item) => ({
      id: `project-${item.uid}`,
      source: item.isMilestone ? "Project milestone" : "Project task",
      title: item.name,
      owner: item.resourceNames?.join(", "),
      stream: item.stream,
      status: item.status,
      priority: item.isCritical ? "Critical" : item.roadmapMilestone ? "Roadmap milestone" : undefined,
      loggedDate: item.startDate,
      dueDate: item.finishDate,
      completionDate: isCompleteStatus(item.status) ? item.finishDate : undefined,
      description: [item.milestoneType, item.approvalBody, item.dependencyLevel].filter(Boolean).join(" · "),
      latestUpdate: item.delayDays && item.delayDays > 0 ? `${item.delayDays} days delayed against baseline` : undefined,
      links: [item.id ? `Project ID ${item.id}` : undefined, item.wbs].filter(Boolean).join(" · "),
      dashboardFlag: item.boardReportable || item.decisionRequired || item.roadmapMilestone,
    }));
  return [...trackerItems, ...projectItems];
}

function csvEscape(value?: string | number | boolean): string {
  const text = value === undefined || value === null ? "" : String(value);
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function downloadTeamActionsCsv(items: TeamWorkItem[], schedule: ProgrammeSchedule) {
  const header = ["Source", "Title", "Owner", "Workstream", "Status", "Priority", "Logged date", "Due date", "Completion date", "Notes"];
  const rows = items.map((item) => [
    item.source,
    item.title,
    item.owner,
    item.stream,
    statusLabel(item),
    item.priority,
    formatDate(item.loggedDate ?? item.meetingDate),
    formatDate(item.dueDate),
    formatDate(item.completionDate),
    item.latestUpdate ?? item.description,
  ]);
  const csv = [header, ...rows].map((row) => row.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${schedule.title.replace(/[^a-z0-9]+/gi, "-").toLowerCase()}-team-actions.csv`;
  link.click();
  URL.revokeObjectURL(link.href);
}

function programmeMilestones(schedule: ProgrammeSchedule): ProgrammeItem[] {
  return schedule.items
    .filter((item) => item.isMilestone || item.roadmapMilestone)
    .sort((a, b) => itemImportance(b) - itemImportance(a) || bySoonest(a.finishDate, b.finishDate));
}

function periodMilestones(schedule: ProgrammeSchedule, window: DateWindow): ProgrammeItem[] {
  return programmeMilestones(schedule)
    .filter((item) => dateWithin(item.finishDate, window) && item.status !== "complete")
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
}

type GanttLevel = "executive" | "milestones" | "all";

const ganttLevels: Array<{ key: GanttLevel; label: string; description: string }> = [
  {
    key: "executive",
    label: "Executive",
    description: "Executive milestones and the key linked predecessor rows, shown in a Microsoft Project-style Gantt.",
  },
  {
    key: "milestones",
    label: "Standard Milestones",
    description: "Standard programme milestones, including roadmap and board-reportable milestones, excluding executive-only outcomes.",
  },
  {
    key: "all",
    label: "All Levels",
    description: "All active dated Project plan rows, preserving the outline level and summary task structure.",
  },
];

function uniqueItems(items: ProgrammeItem[]): ProgrammeItem[] {
  return [...new Map(items.map((item) => [item.uid, item])).values()];
}

function projectOrder(a: ProgrammeItem, b: ProgrammeItem): number {
  const aId = Number(a.id);
  const bId = Number(b.id);
  if (Number.isFinite(aId) && Number.isFinite(bId) && aId !== bId) return aId - bId;
  return (a.outlineNumber ?? a.wbs ?? "").localeCompare(b.outlineNumber ?? b.wbs ?? "", undefined, { numeric: true }) || bySoonest(a.startDate ?? a.finishDate, b.startDate ?? b.finishDate);
}

function ganttItemsForLevel(schedule: ProgrammeSchedule, dateWindow: DateWindow, level: GanttLevel): ProgrammeItem[] {
  const withinWindow = (item: ProgrammeItem) => item.isActive && overlapsDateWindow(item, dateWindow) && (parseDate(item.startDate) || parseDate(item.finishDate));
  if (level === "executive") {
    const model = buildExecutiveRoadmapModel(schedule, undefined, dateWindow);
    return uniqueItems(model.paths.flatMap((path) => [...path.dependencies, path.outcome]))
      .filter(withinWindow)
      .sort(projectOrder);
  }
  if (level === "milestones") {
    return schedule.items
      .filter((item) => withinWindow(item) && (item.isMilestone || item.roadmapMilestone || item.boardReportable) && !item.executiveMilestone)
      .sort(projectOrder);
  }
  return schedule.items
    .filter(withinWindow)
    .sort(projectOrder);
}

function lateMilestones(schedule: ProgrammeSchedule): ProgrammeItem[] {
  return programmeMilestones(schedule)
    .filter((item) => item.status === "late" || Boolean(item.delayDays && item.delayDays > 0))
    .sort((a, b) => (b.delayDays ?? 0) - (a.delayDays ?? 0));
}

type ExecutiveTone = "green" | "blue" | "amber" | "red" | "grey";

type ExecutivePath = {
  outcome: ProgrammeItem;
  dependencies: ProgrammeItem[];
};

type ExecutiveMode = "upcoming" | "delivered" | "other";

type ExecutivePathNode = {
  item: ProgrammeItem;
  depth: number;
};

type TeamWorkItem = {
  id: string;
  source: "Meeting action" | "Project task" | "Project milestone";
  title: string;
  owner?: string;
  stream?: string;
  status?: string;
  priority?: string;
  loggedDate?: string;
  dueDate?: string;
  completionDate?: string;
  meetingDate?: string;
  description?: string;
  latestUpdate?: string;
  links?: string;
  dashboardFlag?: boolean;
};

type TeamStatusFilter = "open" | "due-soon" | "overdue" | "blocked" | "completed";

const teamStatusFilters: Array<{ key: TeamStatusFilter; label: string }> = [
  { key: "open", label: "Open" },
  { key: "due-soon", label: "Due soon" },
  { key: "overdue", label: "Overdue" },
  { key: "blocked", label: "Blocked" },
  { key: "completed", label: "Completed" },
];

const executiveToneLabels: Record<ExecutiveTone, string> = {
  green: "GREEN",
  blue: "PLANNED",
  amber: "AMBER",
  red: "RED",
  grey: "NOT ASSESSED",
};

function normaliseText(value?: string): string {
  return (value ?? "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
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
  if (progress.length) parts.push(`This week progressed ${progress.join(" and ").replace(/\.$/, "")}.`);
  if (priorities.length) parts.push(`Next focus is ${priorities.join(" and ").replace(/\.$/, "")}.`);
  if (rag && blocker) parts.push(`The programme remains ${rag} due to ${blocker.replace(/\.$/, "")}.`);
  else if (blocker) parts.push(`The main blocker is ${blocker.replace(/\.$/, "")}.`);
  return parts.length ? parts.join(" ") : undefined;
}

function significantDependency(item: ProgrammeItem): boolean {
  const level = `${item.milestoneLevel ?? ""} ${item.dependencyLevel ?? ""}`.toLowerCase();
  return Boolean(
    item.executiveMilestone ||
      item.dependencyAnchor ||
      item.governanceGate ||
      item.roadmapMilestone ||
      item.boardReportable ||
      item.decisionRequired ||
      item.externalDependency ||
      level.includes("executive") ||
      level.includes("board") ||
      level.includes("gate"),
  );
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

function importantExecutiveDependency(item: ProgrammeItem): boolean {
  return executiveDependencyScore(item) > 0;
}

function isDeliveredItem(item: ProgrammeItem): boolean {
  return item.status === "complete" || item.percentComplete === 100;
}

function isHighLevelMilestone(item: ProgrammeItem): boolean {
  const level = normaliseText(item.milestoneLevel);
  return Boolean(
    item.executiveMilestone ||
      item.boardReportable ||
      item.roadmapMilestone ||
      level.includes("executive") ||
      level.includes("board"),
  );
}

function isHistoricDeliveredItem(item: ProgrammeItem, window: DateWindow): boolean {
  const finishDate = parseDate(item.finishDate);
  return Boolean(isDeliveredItem(item) && finishDate && window.start && finishDate < window.start);
}

function isForwardLookingExecutiveItem(item: ProgrammeItem, window: DateWindow): boolean {
  if (isHistoricDeliveredItem(item, window)) return false;
  const finishDate = parseDate(item.finishDate);
  if (window.end && finishDate && finishDate > window.end) return false;
  return true;
}

function executiveEnablerScore(item: ProgrammeItem): number {
  const text = normaliseText([
    item.name,
    item.stream,
    item.milestoneType,
    item.dependencyLevel,
    item.approvalBody,
    item.targetMilestone,
  ].filter(Boolean).join(" "));
  let score = executiveDependencyScore(item);
  if (item.isMilestone) score += 20;
  if (item.isCritical) score += 15;
  if (item.externalDependency) score += 10;
  [
    "contract",
    "supplier",
    "commercial",
    "procurement",
    "signed",
    "signature",
    "approval",
    "approved",
    "minister",
    "readiness",
    "ready",
    "enablement",
    "enabled",
    "implementation",
    "transition",
    "registrar employed",
    "companies registry",
  ].forEach((keyword) => {
    if (text.includes(keyword)) score += 35;
  });
  return score;
}

function executiveTone(item?: ProgrammeItem): ExecutiveTone {
  return executiveToneAssessment(item).tone;
}

function recoveryConfidence(weekly: ReturnType<typeof latestWeeklySummary>, outcomes: ProgrammeItem[]): ExecutiveTone {
  const goLive = normaliseText(weekly?.goLiveConfidence);
  const tones = outcomes.map(executiveTone);
  if (goLive.includes("high") || goLive.includes("green")) return "green";
  if (goLive.includes("low") || tones.includes("amber") || normaliseText(weekly?.overallRag).includes("red")) return "amber";
  if (tones.includes("red")) return "red";
  return "blue";
}

function collectPredecessorChain(outcome: ProgrammeItem, byUid: Map<string, ProgrammeItem>): ProgrammeItem[] {
  const found = new Map<string, ProgrammeItem>();
  const seen = new Set<string>();
  const walk = (item: ProgrammeItem, depth: number) => {
    if (depth > 8 || seen.has(item.uid)) return;
    seen.add(item.uid);
    item.predecessors.forEach((link) => {
      if (!link.predecessorUid) return;
      const predecessor = byUid.get(link.predecessorUid);
      if (!predecessor) return;
      found.set(predecessor.uid, predecessor);
      walk(predecessor, depth + 1);
    });
  };
  walk(outcome, 0);
  return [...found.values()];
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
    .map((link) => link.predecessorUid ? byUid.get(link.predecessorUid) : undefined)
    .filter((predecessor): predecessor is ProgrammeItem => Boolean(predecessor?.isActive));
  return [...new Map(predecessors.map((predecessor) => [predecessor.uid, predecessor])).values()]
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
}

function closestDirectPredecessors(item: ProgrammeItem, byUid: Map<string, ProgrammeItem>, limit = 2): ProgrammeItem[] {
  const targetDate = parseDate(item.startDate) ?? parseDate(item.finishDate);
  return directPredecessors(item, byUid)
    .filter((predecessor) => !predecessor.isSummary)
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

function visibleExecutivePathItem(item: ProgrammeItem): boolean {
  return Boolean(item.isActive && !item.isSummary && (item.isMilestone || meaningfulText(item.milestoneLevel)));
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

function collectPredecessorDependencies(outcome: ProgrammeItem, byUid: Map<string, ProgrammeItem>, window: DateWindow): ProgrammeItem[] {
  const chain = collectPredecessorChainWithDepth(outcome, byUid);
  return chain
    .filter((node) => visibleExecutivePathItem(node.item))
    .filter((node) => isForwardLookingExecutiveItem(node.item, window))
    .sort((a, b) => bySoonest(a.item.finishDate, b.item.finishDate) || a.depth - b.depth)
    .map((node) => node.item);
}

function executiveDependencyPaths(schedule: ProgrammeSchedule, window: DateWindow): ExecutivePath[] {
  return buildExecutiveRoadmapModel(schedule, undefined, window).paths;
}

function outlineParts(item: ProgrammeItem): string[] {
  return (item.outlineNumber ?? item.wbs ?? "").split(".").map((part) => part.trim()).filter(Boolean);
}

function itemTopSummary(item: ProgrammeItem, schedule: ProgrammeSchedule): ProgrammeItem | undefined {
  const top = outlineParts(item)[0];
  if (!top) return undefined;
  return schedule.items.find((entry) => entry.isSummary && (entry.outlineNumber ?? entry.wbs) === top);
}

function itemOutlineSection(item: ProgrammeItem, schedule: ProgrammeSchedule): ProgrammeItem | undefined {
  const outline = item.outlineNumber ?? item.wbs ?? "";
  if (!outline) return undefined;
  return schedule.items
    .filter((entry) => entry.isSummary && entry.outlineLevel <= 3)
    .filter((entry) => {
      const prefix = entry.outlineNumber ?? entry.wbs ?? "";
      return Boolean(prefix && outline !== prefix && outline.startsWith(`${prefix}.`));
    })
    .sort((a, b) => b.outlineLevel - a.outlineLevel)[0];
}

function itemGroupLabel(item: ProgrammeItem, schedule: ProgrammeSchedule): string {
  const top = item.stream ?? itemTopSummary(item, schedule)?.name ?? "Unassigned";
  const section = itemOutlineSection(item, schedule);
  if (!section || section.name === top) return top;
  return `${top} / ${section.name}`;
}

function itemSearchText(item: ProgrammeItem, schedule: ProgrammeSchedule): string {
  return normaliseText([
    item.name,
    item.stream,
    itemGroupLabel(item, schedule),
    item.outlineNumber,
    item.wbs,
    item.status,
    item.milestoneLevel,
    item.roadmapView,
    item.resourceNames?.join(" "),
  ].filter(Boolean).join(" "));
}

function relatedTrackerItems<T extends { stream?: string; title: string; dashboardFlag?: boolean }>(items: T[], item?: ProgrammeItem): T[] {
  const stream = normaliseText(item?.stream);
  const target = normaliseText(item?.targetMilestone ?? item?.name);
  return sortFlaggedFirst(items.filter((entry) => {
    const entryStream = normaliseText(entry.stream);
    const entryText = normaliseText(entry.title);
    return Boolean((stream && entryStream === stream) || (target && entryText.includes(target.split(" ").slice(0, 3).join(" "))));
  })).slice(0, 4);
}

function openRisks(tracker?: TrackerData): TrackerRisk[] {
  return tracker?.risks.filter((risk) => isOpenStatus(risk.status)) ?? [];
}

function openIssues(tracker?: TrackerData): TrackerIssue[] {
  return tracker?.issues.filter((issue) => isOpenStatus(issue.status)) ?? [];
}

function openActions(tracker?: TrackerData): TrackerAction[] {
  return tracker?.actions.filter((action) => isOpenStatus(action.status)) ?? [];
}

function openDecisions(tracker?: TrackerData): TrackerDecision[] {
  return tracker?.decisions.filter((decision) => isOpenStatus(decision.status)) ?? [];
}

function ReportingPeriodControl({
  filters,
  setFilters,
}: {
  filters: ProgrammeFilters;
  setFilters: React.Dispatch<React.SetStateAction<ProgrammeFilters>>;
}) {
  const update = <K extends keyof ProgrammeFilters>(key: K, value: ProgrammeFilters[K]) =>
    setFilters((current) => ({ ...current, [key]: value }));
  return (
    <div className="report-period">
      <label className="field">
        <span>Reporting date window</span>
        <select value={filters.datePreset} onChange={(event) => update("datePreset", event.target.value as ProgrammeFilters["datePreset"])}>
          <option value="all">Full programme</option>
          <option value="status-forward">Status date forward</option>
          <option value="current-forward">Current date forward</option>
          <option value="next-30">Next 30 days</option>
          <option value="next-60">Next 60 days</option>
          <option value="next-90">Next 90 days</option>
          <option value="custom">Custom</option>
        </select>
      </label>
      {filters.datePreset === "custom" ? (
        <>
          <label className="field">
            <span>From</span>
            <input type="date" value={filters.dateStart} onChange={(event) => update("dateStart", event.target.value)} />
          </label>
          <label className="field">
            <span>To</span>
            <input type="date" value={filters.dateEnd} onChange={(event) => update("dateEnd", event.target.value)} />
          </label>
        </>
      ) : null}
    </div>
  );
}

function PageIntro({
  title,
  children,
  tracker,
}: {
  title: string;
  children: React.ReactNode;
  tracker?: TrackerData;
}) {
  return (
    <div className="page-intro">
      <div>
        <h2>{title}</h2>
        <p>{children}</p>
      </div>
      <span className={tracker ? "source-pill loaded" : "source-pill"}>
        {tracker ? `Tracker: ${tracker.sourceFileName}` : "Tracker not imported"}
      </span>
    </div>
  );
}

function StatGrid({ cards }: { cards: Array<[string, string, string?]> }) {
  return (
    <section className="report-stat-grid">
      {cards.map(([label, value, tone]) => (
        <div className={`report-stat ${tone ?? ""}`} key={label}>
          <span>{label}</span>
          <strong>{value}</strong>
        </div>
      ))}
    </section>
  );
}

function ItemList({
  title,
  items,
  empty = "No items found.",
}: {
  title: string;
  items: Array<{ id: string; title: string; meta?: string; status?: string }>;
  empty?: string;
}) {
  return (
    <section className="report-card">
      <h3>{title}</h3>
      {items.length ? (
        <div className="report-list">
          {items.slice(0, 8).map((item) => (
            <article key={item.id}>
              <div>
                <strong>{item.title}</strong>
                {item.meta ? <span>{item.meta}</span> : null}
              </div>
              {item.status ? <span className={`status-pill ${item.status.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`}>{item.status}</span> : null}
            </article>
          ))}
        </div>
      ) : (
        <p className="empty-state">{empty}</p>
      )}
    </section>
  );
}

function HomeDashboard({ schedule, tracker, dateWindow }: { schedule: ProgrammeSchedule; tracker?: TrackerData; dateWindow: DateWindow }) {
  const weekly = latestWeeklySummary(tracker);
  const milestones = programmeMilestones(schedule).slice(0, 5);
  const upcoming = periodMilestones(schedule, dateWindow).slice(0, 5);
  const risks = openRisks(tracker).filter((risk) => risk.dashboardFlag || isRedOrAmber(risk.rag)).slice(0, 5);
  const decisions = [
    ...schedule.items.filter((item) => item.decisionRequired || item.governanceGate).map((item) => ({
      id: `plan-${item.uid}`,
      title: item.name,
      meta: `${formatDate(item.finishDate)}${item.approvalBody ? ` · ${item.approvalBody}` : ""}`,
      status: item.status,
    })),
    ...openDecisions(tracker).map((decision) => ({
      id: `decision-${decision.id}`,
      title: decision.title,
      meta: `${formatDate(decision.decisionRequiredBy)}${decision.decisionMaker ? ` · ${decision.decisionMaker}` : ""}`,
      status: decision.status,
    })),
  ].slice(0, 8);

  return (
    <>
      <PageIntro title="Home Dashboard" tracker={tracker}>A single-page programme position combining the imported project plan and the meeting tracker.</PageIntro>
      <StatGrid cards={[
        ["Overall RAG", weekly?.overallRag ?? "Not set", isRedOrAmber(weekly?.overallRag) ? "warn" : ""],
        ["Go live date", formatDate(schedule.finishDate)],
        ["Top risk", risks[0]?.title ?? "None flagged"],
        ["Next milestone", upcoming[0]?.name ?? "None in window"],
      ]} />
      <div className="report-grid two">
        <ItemList title="Top 5 Milestones" items={milestones.map((item) => ({ id: item.uid, title: item.name, meta: `${formatDate(item.finishDate)} · ${item.stream ?? "No stream"}`, status: item.status }))} />
        <ItemList title="Next Window" items={upcoming.map((item) => ({ id: item.uid, title: item.name, meta: `${formatDate(item.finishDate)} · ${item.stream ?? "No stream"}`, status: item.status }))} />
        <ItemList title="Risks Needing Attention" items={risks.map((risk) => ({ id: risk.id, title: risk.title, meta: `${risk.stream ?? "No stream"} · ${risk.owner ?? "No owner"}`, status: risk.rag ?? risk.status }))} />
        <ItemList title="Decisions and Actions Due Soon" items={decisions} />
      </div>
    </>
  );
}

function CEOView({ schedule, tracker }: { schedule: ProgrammeSchedule; tracker?: TrackerData }) {
  const milestones = programmeMilestones(schedule).filter((item) => item.executiveMilestone || itemImportance(item) >= 4).slice(0, 5);
  const decisions = openDecisions(tracker).filter((decision) => decision.dashboardFlag || /ceo|senior|sro|director/i.test(`${decision.decisionMaker} ${decision.owner}`));
  const risks = openRisks(tracker).filter((risk) => risk.dashboardFlag || isRedOrAmber(risk.rag)).slice(0, 5);
  return (
    <>
      <PageIntro title="CEO View" tracker={tracker}>The two-minute view: decisions, blockers and the few milestones that genuinely need senior attention.</PageIntro>
      <StatGrid cards={[
        ["Status", latestWeeklySummary(tracker)?.overallRag ?? "Not set"],
        ["Decision needed", decisions[0]?.title ?? "None flagged"],
        ["Main blocker", risks[0]?.title ?? "None flagged"],
        ["Next milestone", milestones[0]?.name ?? "None flagged"],
      ]} />
      <div className="report-grid two">
        <ItemList title="Top 5 Milestones Timeline" items={milestones.map((item) => ({ id: item.uid, title: item.name, meta: `${formatDate(item.finishDate)} · ${item.stream ?? "No stream"}`, status: item.status }))} />
        <ItemList title="CEO Decisions" items={decisions.map((decision) => ({ id: decision.id, title: decision.title, meta: `${formatDate(decision.decisionRequiredBy)} · ${decision.decisionMaker ?? decision.owner ?? "No owner"}`, status: decision.status }))} />
        <ItemList title="Risks Requiring CEO Attention" items={risks.map((risk) => ({ id: risk.id, title: risk.title, meta: `${risk.stream ?? "No stream"} · ${risk.latestUpdate ?? risk.mitigation ?? ""}`, status: risk.rag ?? risk.status }))} />
        <ItemList title="CEO Actions" items={openActions(tracker).filter((action) => /ceo|senior|sro|director/i.test(`${action.owner} ${action.latestUpdate}`)).map((action) => ({ id: action.id, title: action.title, meta: `${formatDate(action.dueDate)} · ${action.owner ?? "No owner"}`, status: action.status }))} />
      </div>
    </>
  );
}

function ExecutiveSnapshotView({
  schedule,
  tracker,
  dateWindow,
  contextItemUids = [],
  removedItemUids = [],
  laneOrderUids = [],
  onToggleContextItem,
  onRemoveRoadmapItem,
  onRestoreRoadmapItem,
  onReorderLanes,
  onExportSnapshotPdf,
  onExportSnapshotImage,
  onExportSnapshotPosterPdf,
  onExportSnapshotHtml,
}: {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  contextItemUids?: string[];
  removedItemUids?: string[];
  laneOrderUids?: string[];
  onToggleContextItem?: (uid: string) => void;
  onRemoveRoadmapItem?: (uid: string) => void;
  onRestoreRoadmapItem?: (uid: string) => void;
  onReorderLanes?: (orderedUids: string[]) => void;
  onExportSnapshotPdf?: () => void;
  onExportSnapshotImage?: () => void;
  onExportSnapshotPosterPdf?: () => void;
  onExportSnapshotHtml?: () => void;
}) {
  const weekly = latestWeeklySummary(tracker);
  const reportingDate = new Date().toISOString();
  const executiveWindow = { label: dateWindow.label, start: dateWindow.start ?? parseDate(reportingDate) };
  const allExecutiveOutcomes = executiveMilestoneItems(schedule);
  const contextUidSet = useMemo(() => new Set(contextItemUids), [contextItemUids]);
  const baseModel = useMemo(() => buildExecutiveRoadmapModel(schedule, undefined, executiveWindow), [schedule, executiveWindow]);
  const paths = buildExecutiveRoadmapModel(schedule, undefined, executiveWindow, { contextItemUids, removedItemUids, laneOrderUids }).paths
    .filter((path) => isForwardLookingExecutiveItem(path.outcome, executiveWindow));
  const baseDisplayedUids = useMemo(() => new Set(baseModel.paths.flatMap((path) => [path.outcome.uid, ...path.dependencies.map((item) => item.uid)])), [baseModel]);
  const deliveredMilestones = programmeMilestones(schedule)
    .filter((item) => isHighLevelMilestone(item) && isHistoricDeliveredItem(item, executiveWindow))
    .sort((a, b) => (parseDate(b.finishDate)?.getTime() ?? 0) - (parseDate(a.finishDate)?.getTime() ?? 0));
  const deliveredMilestoneUids = useMemo(() => new Set(deliveredMilestones.map((item) => item.uid)), [deliveredMilestones]);
  const [executiveMode, setExecutiveMode] = useState<ExecutiveMode>("upcoming");
  const [selectedUid, setSelectedUid] = useState(paths[0]?.outcome.uid ?? "");
  const [expandedUid, setExpandedUid] = useState("");
  const [draggedLaneUid, setDraggedLaneUid] = useState("");
  const [taskSearch, setTaskSearch] = useState("");
  const [expandedTaskGroups, setExpandedTaskGroups] = useState<Set<string>>(() => new Set());
  const byUid = useMemo(() => new Map(schedule.items.map((item) => [item.uid, item])), [schedule.items]);
  const removedItems = useMemo(() => removedItemUids
    .map((uid) => byUid.get(uid))
    .filter((item): item is ProgrammeItem => Boolean(item))
    .sort(projectOrder), [byUid, removedItemUids]);
  const normalisedTaskSearch = normaliseText(taskSearch);
  const otherProjectTasks = useMemo(() => schedule.items
    .filter((item) => item.isActive && !item.isSummary)
    .filter((item) => !baseDisplayedUids.has(item.uid))
    .filter((item) => !deliveredMilestoneUids.has(item.uid))
    .filter((item) => !normalisedTaskSearch || itemSearchText(item, schedule).includes(normalisedTaskSearch))
    .sort(projectOrder), [baseDisplayedUids, deliveredMilestoneUids, normalisedTaskSearch, schedule]);
  const taskGroups = useMemo(() => {
    const groups = new Map<string, ProgrammeItem[]>();
    otherProjectTasks.forEach((item) => {
      const label = itemGroupLabel(item, schedule);
      groups.set(label, [...(groups.get(label) ?? []), item]);
    });
    return [...groups.entries()].sort(([a], [b]) => a.localeCompare(b));
  }, [otherProjectTasks, schedule]);
  useEffect(() => {
    if (!paths.length) {
      if (selectedUid) setSelectedUid("");
      if (expandedUid) setExpandedUid("");
      return;
    }
    if (!paths.some((path) => path.outcome.uid === selectedUid)) setSelectedUid(paths[0].outcome.uid);
    if (expandedUid && !byUid.has(expandedUid)) setExpandedUid("");
  }, [byUid, expandedUid, paths, selectedUid]);
  const selectedPath = paths.find((path) => path.outcome.uid === selectedUid) ?? paths[0];
  const outcomes = paths.map((path) => path.outcome);
  const deliveryOutcome = programmeDeliveryOutcome(schedule, allExecutiveOutcomes);
  const programmeStatus = meaningfulText(weekly?.overallRag) ?? "Not captured";
  const programmeTone = toneClass(programmeStatus);
  const originalDeliveryDate = deliveryOutcome?.baselineFinish;
  const currentDeliveryDate = deliveryOutcome?.finishDate ?? schedule.finishDate;
  const toggleExpanded = (uid: string) => setExpandedUid((current) => current === uid ? "" : uid);
  const toggleTaskGroup = (group: string) => setExpandedTaskGroups((current) => {
    const next = new Set(current);
    if (next.has(group)) next.delete(group);
    else next.add(group);
    return next;
  });
  const toggleContextItem = (uid: string) => onToggleContextItem?.(uid);
  const removeRoadmapItem = (uid: string) => {
    if (contextUidSet.has(uid)) onToggleContextItem?.(uid);
    else onRemoveRoadmapItem?.(uid);
    if (expandedUid === uid) setExpandedUid("");
  };
  const restoreRoadmapItem = (uid: string) => onRestoreRoadmapItem?.(uid);
  const moveLane = (sourceUid: string, targetUid: string) => {
    if (!onReorderLanes || sourceUid === targetUid) return;
    const currentOrder = paths.map((path) => path.outcome.uid);
    const nextOrder = currentOrder.filter((uid) => uid !== sourceUid);
    const targetIndex = nextOrder.indexOf(targetUid);
    if (targetIndex === -1) return;
    nextOrder.splice(targetIndex, 0, sourceUid);
    onReorderLanes(nextOrder);
  };

  return (
    <>
      <div id="executive-snapshot-export" className="snapshot-export-target">
        <section className="executive-roadmap">
          <header className="exec-roadmap-header">
            <div className="exec-brand">
              <span>DAF</span>
            </div>
            <div>
              <h2>DAF Executive Delivery Roadmap</h2>
            </div>
            <strong>Reporting date: {formatDate(reportingDate)}</strong>
          </header>

          <div className="exec-programme-strip">
            <article className={`exec-programme-status ${programmeTone === "neutral" ? "grey" : programmeTone}`}>
              <span>Programme status</span>
              <strong>{programmeStatus}</strong>
            </article>
            <article>
              <span>Original delivery plan</span>
              <strong>{formatDate(originalDeliveryDate)}</strong>
            </article>
            <article>
              <span>Current forecast</span>
              <strong>{formatDate(currentDeliveryDate)}</strong>
            </article>
            <article>
              <span>Forecast basis</span>
              <strong>{deliveryOutcome?.name ?? "Programme finish"}</strong>
            </article>
          </div>

          <div className="exec-view-tabs" role="tablist" aria-label="Executive roadmap view">
            <button
              className={executiveMode === "upcoming" ? "active" : ""}
              type="button"
              role="tab"
              aria-selected={executiveMode === "upcoming"}
              onClick={() => setExecutiveMode("upcoming")}
            >
              Upcoming roadmap
            </button>
            <button
              className={executiveMode === "delivered" ? "active" : ""}
              type="button"
              role="tab"
              aria-selected={executiveMode === "delivered"}
              onClick={() => setExecutiveMode("delivered")}
            >
              Delivered milestones
            </button>
            <button
              className={executiveMode === "other" ? "active" : ""}
              type="button"
              role="tab"
              aria-selected={executiveMode === "other"}
              onClick={() => setExecutiveMode("other")}
            >
              Other project tasks
            </button>
          </div>

          {executiveMode === "upcoming" ? (
            <>
          <div className="exec-section-label">
            <h3>Executive milestones</h3>
            <p>Only tasks explicitly marked as Executive Milestones in the imported Project plan.</p>
          </div>
          <div className="exec-outcome-cards" aria-label="Executive milestones">
            {outcomes.map((item) => {
              const tone = executiveTone(item);
              return (
                <button
                  className={`exec-outcome-summary ${tone} ${selectedPath?.outcome.uid === item.uid ? "selected" : ""}`}
                  type="button"
                  key={item.uid}
                  onClick={() => setSelectedUid(item.uid)}
                >
                  <span>{formatDate(item.finishDate)}</span>
                  <strong>{item.name}</strong>
                  <em>{executiveToneLabel(item)}</em>
                </button>
              );
            })}
          </div>

          <div className="exec-legend" aria-label="Roadmap legend">
            <span><i className="legend-dot green" /> Complete / confirmed</span>
            <span><i className="legend-dot blue" /> Planned / dated</span>
            <span><i className="legend-dot amber" /> Date assumption / not confirmed</span>
            <span><i className="legend-dot red" /> Blocked / overdue</span>
            <span><i className="legend-dot grey" /> Not assessed</span>
            <span><i className="legend-dot executive" /> Executive dependency</span>
          </div>

          <div className="exec-section-label">
            <h3>Milestone route pathways</h3>
            <p>Key stream milestones plus explicitly linked cross-stream milestone dependencies from the imported Project plan.</p>
          </div>
          <div className="exec-pathways">
            {paths.length ? paths.map((path) => {
              const pathChain = collectPredecessorChain(path.outcome, byUid);
              const expandedItem = path.outcome.uid === selectedPath?.outcome.uid
                ? [path.outcome, ...path.dependencies, ...pathChain].find((item) => item.uid === expandedUid)
                : undefined;
              const expandedPredecessors = expandedItem ? closestDirectPredecessors(expandedItem, byUid, 2) : [];
              const expandedAssessment = expandedItem ? executiveToneAssessment(expandedItem) : undefined;
              return (
              <article
                className={`exec-path ${path.outcome.uid === selectedPath?.outcome.uid ? "selected" : ""} ${draggedLaneUid === path.outcome.uid ? "dragging" : ""}`}
                key={path.outcome.uid}
                onDragOver={(event) => {
                  if (!onReorderLanes || !draggedLaneUid || draggedLaneUid === path.outcome.uid) return;
                  event.preventDefault();
                  event.dataTransfer.dropEffect = "move";
                }}
                onDrop={(event) => {
                  event.preventDefault();
                  const sourceUid = event.dataTransfer.getData("text/plain") || draggedLaneUid;
                  moveLane(sourceUid, path.outcome.uid);
                  setDraggedLaneUid("");
                }}
              >
                <div className="exec-path-title">
                  <h3>{path.outcome.targetMilestone || path.outcome.name}</h3>
                  {onReorderLanes ? (
                    <button
                      className="exec-lane-drag-handle"
                      type="button"
                      draggable
                      onDragStart={(event) => {
                        setDraggedLaneUid(path.outcome.uid);
                        event.dataTransfer.effectAllowed = "move";
                        event.dataTransfer.setData("text/plain", path.outcome.uid);
                      }}
                      onDragEnd={() => setDraggedLaneUid("")}
                      aria-label={`Drag ${path.outcome.name} lane`}
                    >
                      <GripVertical size={16} />
                      <span>Drag lane</span>
                    </button>
                  ) : null}
                </div>
                <div className="exec-path-grid">
                  <div className="exec-sequence" aria-label={`${path.outcome.name} dependency pathway`}>
                    {path.dependencies.length ? path.dependencies.map((item) => {
                      const tone = executiveTone(item);
                      return (
                        <button
                          className={`exec-node ${expandedUid === item.uid ? "expanded" : ""}`}
                          key={item.uid}
                          type="button"
                          onClick={() => {
                            setSelectedUid(path.outcome.uid);
                            toggleExpanded(item.uid);
                          }}
                          aria-expanded={expandedUid === item.uid}
                        >
                          <span>{formatDate(item.finishDate)}</span>
                          <i className={`${tone} ${item.executiveMilestone ? "executive" : ""} ${contextUidSet.has(item.uid) ? "context" : ""}`} />
                          <strong>{item.name}</strong>
                          <small>{contextUidSet.has(item.uid) ? "Added context" : expandedUid === item.uid ? "Hide detail" : "Show detail"}</small>
                        </button>
                      );
                    }) : (
                      <div className="exec-node empty">
                        <span>No links</span>
                        <i className="grey" />
                        <strong>No predecessor links found in the Project plan</strong>
                      </div>
                    )}
                    <ChevronRight className="exec-path-arrow" size={26} />
                  </div>
                  <button
                    className={`exec-outcome-detail ${executiveTone(path.outcome)} ${expandedUid === path.outcome.uid ? "expanded" : ""}`}
                    type="button"
                    onClick={() => {
                      setSelectedUid(path.outcome.uid);
                      toggleExpanded(path.outcome.uid);
                    }}
                    aria-expanded={expandedUid === path.outcome.uid}
                  >
                    <span>{executiveToneLabel(path.outcome)}</span>
                    <strong>{path.outcome.name}</strong>
                    <em>{formatDate(path.outcome.finishDate)}</em>
                    <small>{expandedUid === path.outcome.uid ? "Hide predecessor detail" : "Show predecessor detail"}</small>
                  </button>
                </div>
                {expandedItem ? (
                  <div className="exec-drilldown-panel">
                    <div>
                      <span>Expanded pathway item</span>
                      <h4>{expandedItem.name}</h4>
                      <p>{formatDate(expandedItem.finishDate)} · {expandedItem.stream ?? expandedItem.milestoneLevel ?? expandedItem.dependencyLevel ?? "Project plan item"}</p>
                      {expandedItem.uid !== path.outcome.uid && onRemoveRoadmapItem ? (
                        <button className="exec-context-action danger" type="button" onClick={() => removeRoadmapItem(expandedItem.uid)}>
                          Remove from roadmap
                        </button>
                      ) : null}
                      {expandedAssessment ? (
                        <div className={`exec-rag-explainer ${expandedAssessment.tone}`}>
                          <strong>{expandedItem ? executiveToneLabel(expandedItem) : executiveToneLabels[expandedAssessment.tone]} status rationale</strong>
                          <p>{expandedAssessment.summary}</p>
                          <ul>
                            {expandedAssessment.reasons.map((reason) => (
                              <li key={reason}>{reason}</li>
                            ))}
                          </ul>
                          {expandedAssessment.evidence.length ? (
                            <dl>
                              {expandedAssessment.evidence.map((entry) => (
                                <React.Fragment key={entry.label}>
                                  <dt>{entry.label}</dt>
                                  <dd>{entry.value}</dd>
                                </React.Fragment>
                              ))}
                            </dl>
                          ) : null}
                        </div>
                      ) : null}
                    </div>
                    <div className="exec-drilldown-list">
                      {expandedPredecessors.length ? expandedPredecessors.map((item) => {
                        const tone = executiveTone(item);
                        return (
                          <button
                            className="exec-drilldown-card"
                            type="button"
                            key={item.uid}
                            onClick={() => {
                              setSelectedUid(path.outcome.uid);
                              toggleExpanded(item.uid);
                            }}
                          >
                            <span>{formatDate(item.finishDate)}</span>
                            <strong>{item.name}</strong>
                            <em className={tone}>{executiveToneLabel(item)}</em>
                          </button>
                        );
                      }) : (
                        <p>No direct predecessor items are linked to this item in the Project plan.</p>
                      )}
                    </div>
                  </div>
                ) : null}
              </article>
              );
            }) : (
              <article className="exec-empty-state">
                <h3>No executive milestones found</h3>
                <p>Flag the high-level outcomes in Microsoft Project using the Executive Milestones field, then re-import the XML.</p>
              </article>
            )}
          </div>
          {removedItems.length ? (
            <div className="exec-hidden-roadmap">
              <div>
                <h3>Hidden roadmap items</h3>
                <p>These items are hidden from the Executive Roadmap only. Restore them at any time; the Project plan is unchanged.</p>
              </div>
              <div>
                {removedItems.map((item) => (
                  <button className="exec-hidden-chip" type="button" key={item.uid} onClick={() => restoreRoadmapItem(item.uid)}>
                    <span>{formatDate(item.finishDate)}</span>
                    <strong>{item.name}</strong>
                    <em>Restore</em>
                  </button>
                ))}
              </div>
            </div>
          ) : null}
            </>
          ) : executiveMode === "delivered" ? (
            <>
              <div className="exec-section-label">
                <h3>Delivered milestones</h3>
                <p>Completed high-level programme milestones before {formatDate(executiveWindow.start?.toISOString())}. Add any still-important item back to its roadmap lane.</p>
              </div>
              <div className="exec-delivered-grid">
                {deliveredMilestones.length ? deliveredMilestones.map((item) => {
                  const tone = executiveTone(item);
                  const added = contextUidSet.has(item.uid);
                  return (
                    <article className={`exec-delivered-card ${tone}`} key={item.uid}>
                      <span>{formatDate(item.finishDate)}</span>
                      <strong>{item.name}</strong>
                      <em>{item.stream ?? item.milestoneLevel ?? item.roadmapView ?? "High-level milestone"}</em>
                      {onToggleContextItem ? (
                        <button className={`exec-context-action ${added ? "selected" : ""}`} type="button" onClick={() => toggleContextItem(item.uid)}>
                          {added ? "Remove from lane" : "Add to roadmap lane"}
                        </button>
                      ) : null}
                    </article>
                  );
                }) : (
                  <article className="exec-empty-state">
                    <h3>No delivered milestones found</h3>
                    <p>Completed high-level milestones will appear here once they are marked complete in the imported Project plan.</p>
                  </article>
                )}
              </div>
            </>
          ) : (
            <>
              <div className="exec-section-label">
                <h3>Other project tasks</h3>
                <p>Browse by workstream and outline section, or search, then add useful context to the relevant executive lane.</p>
              </div>
              <label className="exec-task-search">
                <Search size={16} />
                <input
                  type="search"
                  value={taskSearch}
                  onChange={(event) => setTaskSearch(event.target.value)}
                  placeholder="Search task, workstream, owner, status or outline"
                />
              </label>
              <div className="exec-task-browser">
                {taskGroups.length ? taskGroups.map(([group, items]) => {
                  const isExpanded = normalisedTaskSearch ? true : expandedTaskGroups.has(group);
                  const addedCount = items.filter((item) => contextUidSet.has(item.uid)).length;
                  return (
                    <section className="exec-task-group" key={group}>
                      <button className="exec-task-group-header" type="button" onClick={() => toggleTaskGroup(group)} aria-expanded={isExpanded}>
                        <span>{group}</span>
                        <em>{items.length} tasks{addedCount ? ` · ${addedCount} added` : ""}</em>
                      </button>
                      {isExpanded ? (
                        <div className="exec-task-rows">
                          {items.map((item) => {
                            const added = contextUidSet.has(item.uid);
                            return (
                              <article className="exec-task-row" key={item.uid}>
                                <div>
                                  <span>{item.outlineNumber ?? item.wbs ?? "No outline"} · {formatDate(item.finishDate)}</span>
                                  <strong>{item.name}</strong>
                                  <em>{item.isMilestone ? "Milestone" : item.itemType} · {item.status}</em>
                                </div>
                                {onToggleContextItem ? (
                                  <button className={`exec-context-action ${added ? "selected" : ""}`} type="button" onClick={() => toggleContextItem(item.uid)}>
                                    {added ? "Remove" : "Add to lane"}
                                  </button>
                                ) : null}
                              </article>
                            );
                          })}
                        </div>
                      ) : null}
                    </section>
                  );
                }) : (
                  <article className="exec-empty-state">
                    <h3>No matching project tasks</h3>
                    <p>Clear the search or import a fuller Project XML file.</p>
                  </article>
                )}
              </div>
            </>
          )}
        </section>
      </div>
      {onExportSnapshotPdf ? (
        <div className="snapshot-actions">
          <button className="download-action" type="button" onClick={onExportSnapshotPdf}>
            <Download size={15} />
            Table PDF
          </button>
          {onExportSnapshotPosterPdf ? (
            <button className="download-action" type="button" onClick={onExportSnapshotPosterPdf}>
              <Download size={15} />
              Poster PDF
            </button>
          ) : null}
          {onExportSnapshotImage ? (
            <button className="download-action" type="button" onClick={onExportSnapshotImage}>
              <Download size={15} />
              Roadmap PNG
            </button>
          ) : null}
          {onExportSnapshotHtml ? (
            <button className="download-action" type="button" onClick={onExportSnapshotHtml}>
              <Download size={15} />
              HTML
            </button>
          ) : null}
        </div>
      ) : null}
    </>
  );
}

type WeeklyRiskIssueItem = {
  id: string;
  title: string;
  meta?: string;
  status?: string;
  stream?: string;
  dashboardFlag?: boolean;
  kind: "Risk" | "Issue";
};

type WeeklyDragItem = {
  section: WeeklyStatusSectionKey;
  id: string;
};

const weeklySectionLabels: Record<WeeklyStatusSectionKey, string> = {
  milestones: "milestones",
  risksIssues: "risks / issues",
  decisions: "decisions",
  changes: "changes",
};

function weeklyCurationEntry(curation: WeeklyStatusCuration, section: WeeklyStatusSectionKey) {
  return curation[section] ?? { order: [], hidden: [] };
}

function curateWeeklyItems<T>(
  defaultItems: T[],
  allItems: T[],
  section: WeeklyStatusSectionKey,
  curation: WeeklyStatusCuration,
  idFor: (item: T) => string,
  limit: number,
): T[] {
  const entry = weeklyCurationEntry(curation, section);
  const hidden = new Set(entry.hidden);
  const allById = new Map(allItems.map((item) => [idFor(item), item]));
  const ordered = entry.order
    .map((id) => allById.get(id))
    .filter((item): item is T => Boolean(item && !hidden.has(idFor(item))));
  const orderedIds = new Set(ordered.map(idFor));
  const defaults = defaultItems.filter((item) => !hidden.has(idFor(item)) && !orderedIds.has(idFor(item)));
  return [...ordered, ...defaults].slice(0, limit);
}

function remainingWeeklyItems<T>(allItems: T[], visibleItems: T[], idFor: (item: T) => string): T[] {
  const visible = new Set(visibleItems.map(idFor));
  return allItems.filter((item) => !visible.has(idFor(item)));
}

function weeklyRiskIssueCandidates(tracker?: TrackerData): WeeklyRiskIssueItem[] {
  return sortFlaggedFirst([
    ...openRisks(tracker)
      .map((risk) => ({
        id: `risk-${risk.id}`,
        title: risk.title,
        meta: risk.latestUpdate ?? risk.mitigation ?? risk.impact ?? "",
        status: risk.rag ?? risk.status,
        stream: risk.stream,
        dashboardFlag: risk.dashboardFlag,
        kind: "Risk" as const,
      })),
    ...openIssues(tracker)
      .map((issue) => ({
        id: `issue-${issue.id}`,
        title: issue.title,
        meta: issue.latestUpdate ?? issue.requiredAction ?? issue.impact ?? "",
        status: issue.rag ?? issue.priority ?? issue.status,
        stream: issue.stream,
        dashboardFlag: issue.dashboardFlag,
        kind: "Issue" as const,
      })),
  ]);
}

function WeeklyExecutiveStatusView({
  schedule,
  tracker,
  dateWindow,
  curation,
  onUpdateCuration,
  onExportPdf,
}: {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  curation: WeeklyStatusCuration;
  onUpdateCuration: (updater: (current: WeeklyStatusCuration) => WeeklyStatusCuration) => void;
  onExportPdf: () => void;
}) {
  const weekly = latestWeeklySummary(tracker);
  const generatedReportDate = new Date().toISOString();
  const latestWeeklyDate = weekly?.meetingDate ?? weekly?.weekEnding;
  const forwardWindow = { ...dateWindow, start: dateWindow.start ?? parseDate(generatedReportDate) };
  const upcomingMilestoneSource = programmeMilestones(schedule)
    .filter((item) => item.isMilestone && dateWithin(item.finishDate, forwardWindow) && item.status !== "complete")
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
  const completedMilestoneSource = programmeMilestones(schedule)
    .filter((item) => item.isMilestone && isDeliveredItem(item))
    .sort((a, b) => (parseDate(b.finishDate)?.getTime() ?? 0) - (parseDate(a.finishDate)?.getTime() ?? 0));
  const milestoneSource = uniqueItems([...upcomingMilestoneSource, ...completedMilestoneSource]);
  const upcomingMilestones = curateWeeklyItems(upcomingMilestoneSource, milestoneSource, "milestones", curation, (item) => item.uid, 5);
  const allRisksIssues = weeklyRiskIssueCandidates(tracker);
  const defaultRisksIssues = allRisksIssues.filter((item) => item.dashboardFlag || isRedOrAmber(item.status));
  const risksIssues = curateWeeklyItems(defaultRisksIssues, allRisksIssues, "risksIssues", curation, (item) => item.id, 5);
  const allDecisions = (tracker?.decisions ?? [])
    .filter((decision) => !isCompleteStatus(decision.status))
    .sort(decisionSort);
  const defaultDecisions = allDecisions.filter(isOutstandingDecision);
  const decisionsNeeded = curateWeeklyItems(defaultDecisions, allDecisions, "decisions", curation, (decision) => `decision-${decision.id}`, 5);
  const allChanges = (tracker?.changes ?? [])
    .filter((change) => !isCompleteStatus(change.status))
    .sort(changeSort);
  const significantChangeSource = allChanges
    .filter(isSignificantChange)
    .sort(changeSort);
  const significantChanges = curateWeeklyItems(significantChangeSource, allChanges, "changes", curation, (change) => `change-${change.id}`, 5);
  const moreUpcomingMilestones = remainingWeeklyItems(upcomingMilestoneSource, upcomingMilestones, (item) => item.uid);
  const moreCompletedMilestones = remainingWeeklyItems(completedMilestoneSource, upcomingMilestones, (item) => item.uid);
  const moreRisks = remainingWeeklyItems(allRisksIssues.filter((item) => item.kind === "Risk"), risksIssues, (item) => item.id);
  const moreIssues = remainingWeeklyItems(allRisksIssues.filter((item) => item.kind === "Issue"), risksIssues, (item) => item.id);
  const moreDecisions = remainingWeeklyItems(allDecisions, decisionsNeeded, (decision) => `decision-${decision.id}`);
  const moreChanges = remainingWeeklyItems(allChanges, significantChanges, (change) => `change-${change.id}`);
  const [expandedTools, setExpandedTools] = useState<string[]>([]);
  const [showStatusSummaryEditor, setShowStatusSummaryEditor] = useState(true);
  const [dragItem, setDragItem] = useState<WeeklyDragItem | undefined>();
  const ragTone = toneClass(weekly?.overallRag);
  const displayTitle = weeklyProgrammeTitle(schedule.title);
  const deliveryConfidence = meaningfulText(weekly?.goLiveConfidence);
  const forecastToGoLive = forecastToGoLiveLabel(schedule);
  const mainBlocker = meaningfulText(weekly?.mainBlocker) ?? risksIssues[0]?.title;
  const statusSummarySource =
    meaningfulText(weekly?.executiveStatusSummary) ??
    meaningfulText(weekly?.openingLine) ??
    meaningfulText(weekly?.ragRationale) ??
    generatedStatusSummary(weekly, mainBlocker) ??
    "Import the latest tracker to populate the weekly status update.";
  const statusSummary = curation.statusSummaryOverride ?? statusSummarySource;
  const nextMilestone = upcomingMilestoneSource[0];
  const progressItems = [...splitDigest(weekly?.keyProgress, 4), ...splitDigest(weekly?.whatChanged, 2)].slice(0, 5);
  const priorityItems = splitDigest(weekly?.priorityActions, 5);
  const leadershipAsk = meaningfulText(weekly?.askSteerNeeded) ?? meaningfulText(weekly?.decisionsNeeded) ?? decisionsNeeded[0]?.title;
  const trackerStatus = tracker
    ? `${tracker.sourceFileName ?? "Tracker workbook"} | ${tracker.weeklySummaries.length} weekly summaries | ${tracker.risks.length} risks | ${tracker.issues.length} issues | ${tracker.actions.length} actions`
    : "Import the latest tracker workbook to populate RAG, weekly narrative, risks, decisions and actions.";
  const updateCurationSection = (
    section: WeeklyStatusSectionKey,
    transform: (entry: { order: string[]; hidden: string[] }) => { order: string[]; hidden: string[] },
  ) => {
    onUpdateCuration((current) => ({
      ...current,
      [section]: transform(weeklyCurationEntry(current, section)),
    }));
  };
  const removeFromWeeklySection = (section: WeeklyStatusSectionKey, id: string) => {
    updateCurationSection(section, (entry) => ({
      order: entry.order.filter((item) => item !== id),
      hidden: entry.hidden.includes(id) ? entry.hidden : [...entry.hidden, id],
    }));
  };
  const addToWeeklySection = (section: WeeklyStatusSectionKey, id: string) => {
    updateCurationSection(section, (entry) => ({
      order: [id, ...entry.order.filter((item) => item !== id)],
      hidden: entry.hidden.filter((item) => item !== id),
    }));
  };
  const reorderWeeklySection = (section: WeeklyStatusSectionKey, orderedVisibleIds: string[]) => {
    updateCurationSection(section, (entry) => ({
      order: [...orderedVisibleIds, ...entry.order.filter((id) => !orderedVisibleIds.includes(id))],
      hidden: entry.hidden,
    }));
  };
  const moveWeeklyItem = (section: WeeklyStatusSectionKey, sourceId: string, targetId: string, visibleIds: string[]) => {
    if (sourceId === targetId) return;
    const next = visibleIds.filter((id) => id !== sourceId);
    const targetIndex = next.indexOf(targetId);
    if (targetIndex === -1) return;
    next.splice(targetIndex, 0, sourceId);
    reorderWeeklySection(section, next);
  };
  const isToolExpanded = (tool: string) => expandedTools.includes(tool);
  const toggleTool = (tool: string) => setExpandedTools((current) => (
    current.includes(tool) ? current.filter((item) => item !== tool) : [...current, tool]
  ));
  const updateStatusSummary = (value: string | undefined) => {
    onUpdateCuration((current) => ({
      ...current,
      statusSummaryOverride: value,
    }));
  };

  const renderControls = (section: WeeklyStatusSectionKey, id: string, visibleIds: string[]) => (
    <div className="weekly-row-controls">
      <button
        type="button"
        className="weekly-drag-handle"
        draggable
        onDragStart={(event) => {
          setDragItem({ section, id });
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", id);
        }}
        onDragEnd={() => setDragItem(undefined)}
        aria-label={`Drag ${weeklySectionLabels[section]} item`}
      >
        <GripVertical size={14} />
      </button>
      <button type="button" className="weekly-remove-button" onClick={() => removeFromWeeklySection(section, id)} aria-label={`Remove from ${weeklySectionLabels[section]}`}>
        <X size={14} />
      </button>
    </div>
  );

  const rowDropHandlers = (section: WeeklyStatusSectionKey, targetId: string, visibleIds: string[]) => ({
    onDragOver: (event: React.DragEvent) => {
      if (!dragItem || dragItem.section !== section || dragItem.id === targetId) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
    },
    onDrop: (event: React.DragEvent) => {
      event.preventDefault();
      const sourceId = event.dataTransfer.getData("text/plain") || dragItem?.id;
      if (sourceId) moveWeeklyItem(section, sourceId, targetId, visibleIds);
      setDragItem(undefined);
    },
  });

  return (
    <>
      <div id="weekly-status-export" className={`weekly-status weekly-${ragTone}`}>
        <div className={`weekly-source ${tracker && weekly ? "loaded" : "missing"}`}>
          <span>Data source</span>
          <strong>{trackerStatus}</strong>
          <em>{weekly ? `Latest weekly row: ${weekly.id || "untitled"} | ${formatDate(latestWeeklyDate)} | ${weekly.overallRag ?? "RAG not set"}` : "No weekly summary row found"}</em>
        </div>
        <header className="weekly-header">
          <div>
            <span className="snapshot-eyebrow">Weekly executive status</span>
            <h2>{displayTitle}</h2>
            <p>{statusSummary}</p>
          </div>
          <div className={`weekly-rag weekly-rag-${ragTone}`}>
            <span>Overall RAG</span>
            <strong>{meaningfulText(weekly?.overallRag) ?? "Not captured"}</strong>
            <em>{formatDate(generatedReportDate)}</em>
          </div>
        </header>

        {showStatusSummaryEditor ? (
        <section className="weekly-summary-editor">
          <div>
            <span>Status summary</span>
            <p>{curation.statusSummaryOverride !== undefined ? "Edited on dashboard" : meaningfulText(weekly?.executiveStatusSummary) ? "Tracker: Executive Status Summary" : "Auto-generated from latest weekly row"}</p>
          </div>
          <textarea
            value={statusSummary}
            onChange={(event) => updateStatusSummary(event.target.value)}
            rows={3}
            aria-label="Edit executive status summary"
          />
          <div className="weekly-summary-actions">
            {curation.statusSummaryOverride !== undefined ? (
              <button type="button" className="download-action secondary" onClick={() => updateStatusSummary(undefined)}>
                Use tracker summary
              </button>
            ) : null}
            <button type="button" className="download-action secondary" onClick={() => setShowStatusSummaryEditor(false)}>
              Hide editor
            </button>
          </div>
        </section>
        ) : null}

        <section className="weekly-kpis">
          <article>
            <span>Delivery confidence</span>
            <strong>{deliveryConfidence ?? "Not captured"}</strong>
          </article>
          <article>
            <span>Forecast to go live</span>
            <strong>{forecastToGoLive}</strong>
          </article>
          <article>
            <span>Main blocker</span>
            <strong>{mainBlocker ?? "None flagged"}</strong>
          </article>
          <article>
            <span>Next milestone</span>
            <strong>{nextMilestone ? `${formatDate(nextMilestone.finishDate)} - ${nextMilestone.name}` : "None in window"}</strong>
          </article>
        </section>

        <section className="weekly-focus">
          <article className="weekly-panel">
            <h3>Last week</h3>
            {progressItems.map((item) => <p key={item}>{item}</p>)}
            {!progressItems.length ? <p>No weekly progress summary found.</p> : null}
          </article>
          <article className="weekly-panel">
            <h3>This week / next</h3>
            {priorityItems.map((item) => <p key={item}>{item}</p>)}
            {!priorityItems.length ? <p>No priority actions summary found.</p> : null}
          </article>
          <article className="weekly-panel weekly-ask">
            <h3>Leadership ask</h3>
            <p>{leadershipAsk ?? "No current leadership ask flagged."}</p>
          </article>
        </section>

        <section className="weekly-grid">
          <article className="weekly-card">
            <h3>Upcoming milestones</h3>
            {(() => {
              const visibleIds = upcomingMilestones.map((item) => item.uid);
              return upcomingMilestones.map((item) => (
                <div className={`weekly-row curated ${dragItem?.id === item.uid ? "dragging" : ""}`} key={item.uid} {...rowDropHandlers("milestones", item.uid, visibleIds)}>
                  <div>
                    <span>{formatDate(item.finishDate)}</span>
                    <strong>{item.name}</strong>
                    <em>{item.stream ?? item.milestoneLevel ?? "Milestone"}</em>
                    {isDeliveredItem(item) ? <mark>Completed</mark> : null}
                  </div>
                  {renderControls("milestones", item.uid, visibleIds)}
                </div>
              ));
            })()}
            {!upcomingMilestones.length ? <p>No upcoming milestones found in the selected date window.</p> : null}
          </article>
          <article className="weekly-card">
            <h3>Risks / issues</h3>
            {(() => {
              const visibleIds = risksIssues.map((item) => item.id);
              return risksIssues.map((item) => (
                <div className={`weekly-row curated ${dragItem?.id === item.id ? "dragging" : ""}`} key={item.id} {...rowDropHandlers("risksIssues", item.id, visibleIds)}>
                  <div>
                    <span>{item.status ?? item.kind}</span>
                    <strong>{item.title}</strong>
                    <em>{item.stream ?? item.meta}</em>
                  </div>
                  {renderControls("risksIssues", item.id, visibleIds)}
                </div>
              ));
            })()}
            {!risksIssues.length ? <p>No dashboard or red/amber risks or issues currently flagged.</p> : null}
          </article>
          <article className="weekly-card">
            <h3>Decisions needed</h3>
            {(() => {
              const visibleIds = decisionsNeeded.map((decision) => `decision-${decision.id}`);
              return decisionsNeeded.map((decision) => {
                const id = `decision-${decision.id}`;
                return (
                  <div className={`weekly-row curated ${dragItem?.id === id ? "dragging" : ""}`} key={decision.id} {...rowDropHandlers("decisions", id, visibleIds)}>
                    <div>
                      <span>{formatDateOrText(decision.decisionRequiredBy ?? decision.decisionDate, "Decision date tbc")}</span>
                      <strong>{decision.title}</strong>
                      <em>{decision.decisionMaker ?? decision.owner ?? decision.status ?? "Decision required"}</em>
                    </div>
                    {renderControls("decisions", id, visibleIds)}
                  </div>
                );
              });
            })()}
            {!decisionsNeeded.length ? <p>No outstanding executive decisions currently flagged.</p> : null}
          </article>
          <article className="weekly-card">
            <h3>Significant changes</h3>
            {(() => {
              const visibleIds = significantChanges.map((change) => `change-${change.id}`);
              return significantChanges.map((change) => {
                const id = `change-${change.id}`;
                return (
                  <div className={`weekly-row curated ${dragItem?.id === id ? "dragging" : ""}`} key={change.id} {...rowDropHandlers("changes", id, visibleIds)}>
                    <div>
                      <span>{formatDate(change.lastDiscussedDate ?? change.dateRaised)}</span>
                      <strong>{change.title}</strong>
                      <em>{meaningfulText(change.decisionRequired) ?? meaningfulText(change.impactOnTime) ?? meaningfulText(change.impactOnScope) ?? meaningfulText(change.impactOnCost) ?? meaningfulText(change.impactOnQualityOrBenefits) ?? change.status ?? "Significant change"}</em>
                    </div>
                    {renderControls("changes", id, visibleIds)}
                  </div>
                );
              });
            })()}
            {!significantChanges.length ? <p>No significant changes currently flagged for leadership visibility.</p> : null}
          </article>
        </section>
      </div>
      <section className="weekly-curation-tools" aria-label="Weekly status extra items">
        <div className="weekly-tool-header">
          <div>
            <h3>Status update source items</h3>
            <p>Use these screen-only panels to add relevant items into the weekly status. They are not printed unless you add the item to a main card above.</p>
          </div>
          <button type="button" className="download-action secondary" onClick={() => onUpdateCuration(() => ({}))}>Reset weekly selection</button>
        </div>
        <div className="weekly-tool-buttons">
          {!showStatusSummaryEditor ? <button type="button" onClick={() => setShowStatusSummaryEditor(true)}>Show status editor</button> : null}
          <button type="button" className={isToolExpanded("upcoming") ? "active" : ""} aria-expanded={isToolExpanded("upcoming")} onClick={() => toggleTool("upcoming")}>More upcoming milestones ({moreUpcomingMilestones.length})</button>
          <button type="button" className={isToolExpanded("completed") ? "active" : ""} aria-expanded={isToolExpanded("completed")} onClick={() => toggleTool("completed")}>Completed milestones ({moreCompletedMilestones.length})</button>
          <button type="button" className={isToolExpanded("risks") ? "active" : ""} aria-expanded={isToolExpanded("risks")} onClick={() => toggleTool("risks")}>More risks ({moreRisks.length})</button>
          <button type="button" className={isToolExpanded("issues") ? "active" : ""} aria-expanded={isToolExpanded("issues")} onClick={() => toggleTool("issues")}>More issues ({moreIssues.length})</button>
          <button type="button" className={isToolExpanded("decisions") ? "active" : ""} aria-expanded={isToolExpanded("decisions")} onClick={() => toggleTool("decisions")}>More decisions ({moreDecisions.length})</button>
          <button type="button" className={isToolExpanded("changes") ? "active" : ""} aria-expanded={isToolExpanded("changes")} onClick={() => toggleTool("changes")}>More changes ({moreChanges.length})</button>
        </div>
        {isToolExpanded("upcoming") ? (
          <WeeklySourceList
            title="More upcoming milestones"
            items={moreUpcomingMilestones.map((item) => ({ id: item.uid, title: item.name, eyebrow: formatDate(item.finishDate), meta: item.stream ?? item.milestoneLevel ?? "Milestone" }))}
            onAdd={(id) => addToWeeklySection("milestones", id)}
          />
        ) : null}
        {isToolExpanded("completed") ? (
          <WeeklySourceList
            title="Completed milestones"
            items={moreCompletedMilestones.map((item) => ({ id: item.uid, title: item.name, eyebrow: formatDate(item.finishDate), meta: item.stream ?? item.milestoneLevel ?? "Milestone" }))}
            onAdd={(id) => addToWeeklySection("milestones", id)}
          />
        ) : null}
        {isToolExpanded("risks") ? (
          <WeeklySourceList
            title="More risks"
            items={moreRisks.map((item) => ({ id: item.id, title: item.title, eyebrow: item.status ?? "Risk", meta: item.stream ?? item.meta ?? "" }))}
            onAdd={(id) => addToWeeklySection("risksIssues", id)}
          />
        ) : null}
        {isToolExpanded("issues") ? (
          <WeeklySourceList
            title="More issues"
            items={moreIssues.map((item) => ({ id: item.id, title: item.title, eyebrow: item.status ?? "Issue", meta: item.stream ?? item.meta ?? "" }))}
            onAdd={(id) => addToWeeklySection("risksIssues", id)}
          />
        ) : null}
        {isToolExpanded("decisions") ? (
          <WeeklySourceList
            title="More decisions"
            items={moreDecisions.map((decision) => ({ id: `decision-${decision.id}`, title: decision.title, eyebrow: formatDateOrText(decision.decisionRequiredBy ?? decision.decisionDate, "Decision date tbc"), meta: decision.decisionMaker ?? decision.owner ?? decision.status ?? "" }))}
            onAdd={(id) => addToWeeklySection("decisions", id)}
          />
        ) : null}
        {isToolExpanded("changes") ? (
          <WeeklySourceList
            title="More changes"
            items={moreChanges.map((change) => ({ id: `change-${change.id}`, title: change.title, eyebrow: formatDate(change.lastDiscussedDate ?? change.dateRaised), meta: meaningfulText(change.decisionRequired) ?? meaningfulText(change.impactOnTime) ?? meaningfulText(change.impactOnScope) ?? change.status ?? "" }))}
            onAdd={(id) => addToWeeklySection("changes", id)}
          />
        ) : null}
      </section>
      <div className="snapshot-actions">
        <button className="download-action" type="button" onClick={onExportPdf}>
          <Download size={15} />
          Download A4 Weekly Status PDF
        </button>
      </div>
    </>
  );
}

function WeeklySourceList({
  title,
  items,
  onAdd,
}: {
  title: string;
  items: Array<{ id: string; title: string; eyebrow?: string; meta?: string }>;
  onAdd: (id: string) => void;
}) {
  return (
    <section className="weekly-source-list">
      <h4>{title}</h4>
      {items.length ? (
        <div>
          {items.map((item) => (
            <article className="weekly-source-row" key={item.id}>
              <div>
                <span>{item.eyebrow ?? "Available"}</span>
                <strong>{item.title}</strong>
                {item.meta ? <em>{item.meta}</em> : null}
              </div>
              <button type="button" onClick={() => onAdd(item.id)}>Add to status</button>
            </article>
          ))}
        </div>
      ) : (
        <p>No additional items available.</p>
      )}
    </section>
  );
}

function BoardReportView({ schedule, tracker, dateWindow }: { schedule: ProgrammeSchedule; tracker?: TrackerData; dateWindow: DateWindow }) {
  const weekly = latestWeeklySummary(tracker);
  const weeklyByDate = tracker?.weeklySummaries
    .slice()
    .sort((a, b) => (parseDate(b.weekEnding)?.getTime() ?? 0) - (parseDate(a.weekEnding)?.getTime() ?? 0));
  const boardMilestones = programmeMilestones(schedule).filter((item) => item.boardReportable || /board/i.test(item.milestoneLevel ?? "")).slice(0, 8);
  const nextPeriod = periodMilestones(schedule, dateWindow).slice(0, 8);
  return (
    <>
      <PageIntro title="Board Report View" tracker={tracker}>Formal reporting content for board packs, escalation and the next reporting period.</PageIntro>
      <StatGrid cards={[
        ["Current RAG", weekly?.overallRag ?? "Not set"],
        ["Previous RAG", weeklyByDate?.[1]?.overallRag ?? "Not set"],
        ["Movement", weekly?.whatChanged ? "Updated" : "Not set"],
        ["Decision required", openDecisions(tracker).length.toString()],
      ]} />
      <div className="report-grid two">
        <ItemList title="Board Reportable Milestones" items={boardMilestones.map((item) => ({ id: item.uid, title: item.name, meta: `${formatDate(item.finishDate)} · ${item.stream ?? "No stream"}`, status: item.status }))} />
        <ItemList title="Risks for Escalation" items={openRisks(tracker).filter((risk) => risk.dashboardFlag || isRedOrAmber(risk.rag)).map((risk) => ({ id: risk.id, title: risk.title, meta: risk.latestUpdate ?? risk.mitigation, status: risk.rag ?? risk.status }))} />
        <ItemList title="Decisions Required" items={openDecisions(tracker).map((decision) => ({ id: decision.id, title: decision.title, meta: `${formatDate(decision.decisionRequiredBy)} · ${decision.decisionMaker ?? "No decision maker"}`, status: decision.status }))} />
        <ItemList title="Next Reporting Period" items={nextPeriod.map((item) => ({ id: item.uid, title: item.name, meta: `${formatDate(item.finishDate)} · ${item.approvalBody ?? item.stream ?? "No stream"}`, status: item.status }))} />
      </div>
    </>
  );
}

function ReportingRoadmapView({ schedule, dateWindow, selected, onSelect }: { schedule: ProgrammeSchedule; dateWindow: DateWindow; selected?: ProgrammeItem; onSelect: (item: ProgrammeItem) => void }) {
  const items = programmeMilestones(schedule).filter((item) => item.roadmapMilestone || itemImportance(item) >= 3);
  return (
    <>
      <PageIntro title="Reporting Roadmap">An audience-friendly roadmap drawn from roadmap, executive and board-level milestones.</PageIntro>
      <Timeline schedule={schedule} items={items} selected={selected} onSelect={onSelect} dateWindow={dateWindow} />
    </>
  );
}

function ganttTone(item: ProgrammeItem): "summary" | "complete" | "risk" | "late" | "task" {
  if (item.status === "late" || item.status === "blocked" || (item.delayDays ?? 0) > 0) return "late";
  if (item.status === "at-risk" || item.externalDependency || item.decisionRequired) return "risk";
  if (item.status === "complete") return "complete";
  if (item.isSummary) return "summary";
  return "task";
}

function ganttScaleTicks(bounds: ReturnType<typeof timelineBounds>) {
  const ticks: Array<{ label: string; left: number }> = [];
  const cursor = new Date(Date.UTC(bounds.min.getUTCFullYear(), bounds.min.getUTCMonth(), 1));
  while (cursor <= bounds.max) {
    ticks.push({
      label: new Intl.DateTimeFormat("en-GB", { month: "short", year: "2-digit" }).format(cursor),
      left: clamp(((cursor.getTime() - bounds.min.getTime()) / bounds.span) * 100, 0, 100),
    });
    cursor.setUTCMonth(cursor.getUTCMonth() + 1);
  }
  return ticks;
}

function immediatePredecessors(item: ProgrammeItem | undefined, byUid: Map<string, ProgrammeItem>): ProgrammeItem[] {
  if (!item) return [];
  return uniqueItems(item.predecessors
    .map((link) => link.predecessorUid ? byUid.get(link.predecessorUid) : undefined)
    .filter((entry): entry is ProgrammeItem => Boolean(entry?.isActive)))
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
}

function predecessorText(item: ProgrammeItem, byUid: Map<string, ProgrammeItem>): string {
  const values = item.predecessors
    .map((link) => {
      if (!link.predecessorUid) return undefined;
      const predecessor = byUid.get(link.predecessorUid);
      return predecessor?.id ?? predecessor?.outlineNumber ?? predecessor?.wbs ?? predecessor?.uid;
    })
    .filter(Boolean);
  return values.length ? values.join(", ") : "-";
}

const GANTT_ROW_HEIGHT = 38;
const MAX_GANTT_ARROWS = 320;

type GanttLinkPath = {
  id: string;
  fromUid: string;
  toUid: string;
  path: string;
};

function ganttDependencyPaths(visibleItems: ProgrammeItem[], byUid: Map<string, ProgrammeItem>, bounds: ReturnType<typeof timelineBounds>): GanttLinkPath[] {
  const visibleIndex = new Map(visibleItems.map((item, index) => [item.uid, index]));
  const paths: GanttLinkPath[] = [];

  visibleItems.forEach((item, toIndex) => {
    item.predecessors.forEach((link) => {
      if (!link.predecessorUid) return;
      const predecessor = byUid.get(link.predecessorUid);
      const fromIndex = predecessor ? visibleIndex.get(predecessor.uid) : undefined;
      if (!predecessor || fromIndex === undefined) return;

      const fromX = positionFor(predecessor.finishDate ?? predecessor.startDate, bounds);
      const toX = positionFor(item.isMilestone ? item.finishDate : item.startDate ?? item.finishDate, bounds);
      const fromY = fromIndex * GANTT_ROW_HEIGHT + (GANTT_ROW_HEIGHT / 2);
      const toY = toIndex * GANTT_ROW_HEIGHT + (GANTT_ROW_HEIGHT / 2);
      const bendX = clamp(Math.max(fromX, toX) + 2.5, 1.5, 98.5);
      const targetX = clamp(toX - 0.7, 0.8, 99.2);

      paths.push({
        id: `${predecessor.uid}-${item.uid}-${paths.length}`,
        fromUid: predecessor.uid,
        toUid: item.uid,
        path: `M ${fromX.toFixed(2)} ${fromY.toFixed(1)} L ${bendX.toFixed(2)} ${fromY.toFixed(1)} L ${bendX.toFixed(2)} ${toY.toFixed(1)} L ${targetX.toFixed(2)} ${toY.toFixed(1)}`,
      });
    });
  });

  return paths;
}

function GanttChart({
  schedule,
  items,
  selected,
  onSelect,
  dateWindow,
  showArrows,
}: {
  schedule: ProgrammeSchedule;
  items: ProgrammeItem[];
  selected?: ProgrammeItem;
  onSelect: (item: ProgrammeItem) => void;
  dateWindow: DateWindow;
  showArrows: boolean;
}) {
  const visibleItems = items.slice(0, 220);
  const bounds = useMemo(() => timelineBounds(visibleItems.length ? visibleItems : schedule.items, schedule, dateWindow), [dateWindow, schedule, visibleItems]);
  const ticks = useMemo(() => ganttScaleTicks(bounds), [bounds]);
  const statusMarker = positionFor(schedule.statusDate, bounds);
  const byUid = useMemo(() => new Map(schedule.items.map((item) => [item.uid, item])), [schedule.items]);
  const selectedPredecessors = immediatePredecessors(selected, byUid);
  const predecessorIds = new Set(selectedPredecessors.map((item) => item.uid));
  const allDependencyPaths = useMemo(() => ganttDependencyPaths(visibleItems, byUid, bounds), [bounds, byUid, visibleItems]);
  const dependencyPaths = allDependencyPaths.slice(0, MAX_GANTT_ARROWS);
  const hiddenArrowCount = Math.max(0, allDependencyPaths.length - dependencyPaths.length);
  const chartHeight = Math.max(GANTT_ROW_HEIGHT, visibleItems.length * GANTT_ROW_HEIGHT);
  return (
    <section className="gantt-shell">
      <div className="gantt-summary">
        <strong>{visibleItems.length}</strong> items shown
        {items.length > visibleItems.length ? <span> from {items.length} dated items</span> : null}
        <span>{dateWindow.label}</span>
        {showArrows ? <span>{dependencyPaths.length} dependency arrows{hiddenArrowCount ? ` shown, ${hiddenArrowCount} hidden for readability` : ""}</span> : null}
      </div>
      <div className="gantt-legend">
        <span><i className="gantt-key summary" />Summary</span>
        <span><i className="gantt-key task" />Task</span>
        <span><i className="gantt-key complete" />Complete</span>
        <span><i className="gantt-key risk" />At risk</span>
        <span><i className="gantt-key late" />Late / blocked</span>
        <span><Diamond size={13} />Milestone</span>
        <span><i className="gantt-key baseline" />Baseline finish</span>
        <span><i className="gantt-key arrow" />Dependency arrow</span>
        <span><i className="gantt-key dependency" />Selected predecessor</span>
      </div>
      <div className="msp-gantt" style={{ ["--status-left" as string]: `${statusMarker}%` }}>
        <div className="msp-head">ID</div>
        <div className="msp-head task-name">Task Name</div>
        <div className="msp-head">Start</div>
        <div className="msp-head">Finish</div>
        <div className="msp-head">Duration</div>
        <div className="msp-head">% Complete</div>
        <div className="msp-head">Predecessors</div>
        <div className="msp-scale">
          {ticks.map((tick) => (
            <span key={`${tick.label}-${tick.left}`} style={{ left: `${tick.left}%` }}>{tick.label}</span>
          ))}
        </div>
        {visibleItems.map((item) => {
          const start = item.isMilestone ? positionFor(item.finishDate, bounds) : positionFor(item.startDate ?? item.finishDate, bounds);
          const finish = positionFor(item.finishDate ?? item.startDate, bounds);
          const left = Math.min(start, finish);
          const width = Math.max(0.9, Math.abs(finish - start));
          const baseline = positionFor(item.baselineFinish, bounds);
          const tone = ganttTone(item);
          const isSelected = selected?.uid === item.uid;
          const isPredecessor = predecessorIds.has(item.uid);
          const progress = clamp(item.percentComplete ?? 0, 0, 100);
          const rowClass = `${isSelected ? "selected" : ""} ${isPredecessor ? "dependency-predecessor" : ""} ${item.isSummary ? "summary" : ""}`;
          return (
            <React.Fragment key={item.uid}>
              <button
                className={`msp-cell msp-id ${rowClass}`}
                type="button"
                onClick={() => onSelect(item)}
              >
                {item.id ?? item.outlineNumber ?? "-"}
              </button>
              <button
                className={`msp-cell msp-task-name ${rowClass}`}
                type="button"
                onClick={() => onSelect(item)}
                style={{ ["--indent" as string]: `${Math.max(0, item.outlineLevel - 1) * 14}px` }}
              >
                <span>{item.outlineNumber ?? item.wbs ?? ""}</span>
                <strong>{item.name}</strong>
              </button>
              <button className={`msp-cell ${rowClass}`} type="button" onClick={() => onSelect(item)}>{formatDate(item.startDate)}</button>
              <button className={`msp-cell ${rowClass}`} type="button" onClick={() => onSelect(item)}>{formatDate(item.finishDate)}</button>
              <button className={`msp-cell ${rowClass}`} type="button" onClick={() => onSelect(item)}>{durationLabel(item.duration)}</button>
              <button className={`msp-cell ${rowClass}`} type="button" onClick={() => onSelect(item)}>{item.percentComplete ?? 0}%</button>
              <button className={`msp-cell ${rowClass}`} type="button" onClick={() => onSelect(item)}>{predecessorText(item, byUid)}</button>
              <div className={`msp-chart-row ${rowClass}`}>
                {item.baselineFinish ? <i className="gantt-baseline" style={{ left: `${baseline}%` }} /> : null}
                <button
                  className={`gantt-bar ${tone} ${item.isMilestone ? "milestone" : ""} ${item.executiveMilestone ? "executive" : ""}`}
                  type="button"
                  style={{ left: `${left}%`, width: item.isMilestone ? undefined : `${width}%` }}
                  onClick={() => onSelect(item)}
                  title={`${item.name} · ${formatDate(item.startDate)} to ${formatDate(item.finishDate)}`}
                >
                  {!item.isMilestone ? <i className="gantt-progress" style={{ width: `${progress}%` }} /> : null}
                  <span>{item.isMilestone ? formatDate(item.finishDate) : `${formatDate(item.startDate)} - ${formatDate(item.finishDate)}`}</span>
                </button>
              </div>
            </React.Fragment>
          );
        })}
        {showArrows && dependencyPaths.length ? (
          <svg
            className="gantt-link-layer"
            style={{ gridColumn: 8, gridRow: `2 / span ${Math.max(1, visibleItems.length)}` }}
            viewBox={`0 0 100 ${chartHeight}`}
            preserveAspectRatio="none"
            aria-hidden="true"
          >
            <defs>
              <marker id="gantt-arrow-head" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="5" markerHeight="5" orient="auto-start-reverse">
                <path d="M 0 0 L 8 4 L 0 8 z" fill="#2f6380" />
              </marker>
              <marker id="gantt-arrow-head-selected" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="5" markerHeight="5" orient="auto-start-reverse">
                <path d="M 0 0 L 8 4 L 0 8 z" fill="#8f5c00" />
              </marker>
            </defs>
            {dependencyPaths.map((link) => {
              const isSelectedLink = selected?.uid === link.fromUid || selected?.uid === link.toUid;
              return (
                <path
                  key={link.id}
                  className={isSelectedLink ? "selected" : undefined}
                  d={link.path}
                  markerEnd={isSelectedLink ? "url(#gantt-arrow-head-selected)" : "url(#gantt-arrow-head)"}
                />
              );
            })}
          </svg>
        ) : null}
        {!visibleItems.length ? (
          <div className="gantt-empty">No dated Project rows found for this Gantt level and date window.</div>
        ) : null}
      </div>
    </section>
  );
}

function GanttView({ schedule, dateWindow, selected, onSelect }: { schedule: ProgrammeSchedule; dateWindow: DateWindow; selected?: ProgrammeItem; onSelect: (item: ProgrammeItem) => void }) {
  const [level, setLevel] = useState<GanttLevel>("executive");
  const [exportError, setExportError] = useState<string>();
  const [showArrows, setShowArrows] = useState(true);
  const sections = useMemo(() => ganttLevels.map((entry) => ({
    key: entry.key,
    label: entry.label,
    description: entry.description,
    items: ganttItemsForLevel(schedule, dateWindow, entry.key),
  })), [dateWindow, schedule]);
  const active = sections.find((entry) => entry.key === level) ?? sections[0];
  const exportSections = async (targetSections: GanttPdfSection[]) => {
    setExportError(undefined);
    try {
      await exportGanttPdf(schedule, targetSections, dateWindow);
    } catch (err) {
      setExportError(err instanceof Error ? err.message : "The Gantt PDF could not be generated.");
    }
  };

  return (
    <>
      <PageIntro title="Gantt View">A Microsoft Project-style Gantt with task table columns, timeline bars and predecessor references from the imported XML.</PageIntro>
      <section className="gantt-controls">
        <div className="tabs">
          {sections.map((entry) => (
            <button
              className={entry.key === level ? "active" : ""}
              type="button"
              key={entry.key}
              onClick={() => setLevel(entry.key)}
            >
              {entry.label}
            </button>
          ))}
        </div>
        <div className="gantt-downloads">
          <label className="check gantt-arrow-toggle">
            <input type="checkbox" checked={showArrows} onChange={(event) => setShowArrows(event.target.checked)} />
            Show dependency arrows
          </label>
          <button className="download-action" type="button" onClick={() => exportSections([{ label: active.label, items: active.items }])}>
            <Download size={15} />
            Download current level
          </button>
          <button className="download-action secondary" type="button" onClick={() => exportSections(sections.map((entry) => ({ label: entry.label, items: entry.items })))}>
            <Download size={15} />
            Download all levels
          </button>
        </div>
      </section>
      {exportError ? <div className="error">{exportError}</div> : null}
      <section className="gantt-level-note">
        <strong>{active.label}</strong>
        <span>{active.description}</span>
      </section>
      <GanttChart schedule={schedule} items={active.items} selected={selected} onSelect={onSelect} dateWindow={dateWindow} showArrows={showArrows} />
    </>
  );
}

function RisksIssuesView({ tracker }: { tracker?: TrackerData }) {
  const risks = openRisks(tracker);
  const issues = openIssues(tracker);
  return (
    <>
      <PageIntro title="Risks and Issues" tracker={tracker}>Threats, current problems, escalations and the latest tracker updates.</PageIntro>
      <StatGrid cards={[
        ["Open risks", risks.length.toString()],
        ["Red / amber risks", risks.filter((risk) => isRedOrAmber(risk.rag)).length.toString(), "warn"],
        ["Open issues", issues.length.toString()],
        ["Escalated items", [...risks, ...issues].filter((item) => item.dashboardFlag).length.toString()],
      ]} />
      <div className="report-grid two">
        <ItemList title="Top Risks" items={risks.filter((risk) => risk.dashboardFlag || isRedOrAmber(risk.rag)).map((risk) => ({ id: risk.id, title: risk.title, meta: `${risk.stream ?? "No stream"} · ${risk.mitigation ?? risk.latestUpdate ?? ""}`, status: risk.rag ?? risk.status }))} />
        <ItemList title="Open Issues" items={issues.map((issue) => ({ id: issue.id, title: issue.title, meta: `${issue.stream ?? "No stream"} · ${issue.requiredAction ?? issue.latestUpdate ?? ""}`, status: issue.rag ?? issue.status }))} />
      </div>
    </>
  );
}

function ActionsDecisionsView({ schedule, tracker, dateWindow }: { schedule: ProgrammeSchedule; tracker?: TrackerData; dateWindow: DateWindow }) {
  const actions = openActions(tracker).filter((action) => dateWithin(action.dueDate, dateWindow));
  const decisions = openDecisions(tracker).filter((decision) => dateWithin(decision.decisionRequiredBy ?? decision.decisionDate, dateWindow));
  const approvalGates = schedule.items.filter((item) => item.approvalBody || item.governanceGate || item.decisionRequired);
  return (
    <>
      <PageIntro title="Actions and Decisions" tracker={tracker}>Who needs to do what, by when, and which formal approvals are coming from the plan.</PageIntro>
      <div className="report-grid two">
        <ItemList title="Actions Due in Window" items={actions.sort((a, b) => bySoonest(a.dueDate, b.dueDate)).map((action) => ({ id: action.id, title: action.title, meta: `${formatDate(action.dueDate)} · ${action.owner ?? "No owner"}`, status: action.status }))} />
        <ItemList title="Decisions Needed" items={decisions.sort((a, b) => bySoonest(a.decisionRequiredBy, b.decisionRequiredBy)).map((decision) => ({ id: decision.id, title: decision.title, meta: `${formatDate(decision.decisionRequiredBy)} · ${decision.decisionMaker ?? decision.owner ?? "No owner"}`, status: decision.status }))} />
        <ItemList title="Approval Gates from Plan" items={approvalGates.slice(0, 8).map((item) => ({ id: item.uid, title: item.name, meta: `${formatDate(item.finishDate)} · ${item.approvalBody ?? "Governance gate"}`, status: item.status }))} />
        <ItemList title="This Week's Priority Actions" items={openActions(tracker).filter((action) => action.dashboardFlag).map((action) => ({ id: action.id, title: action.title, meta: `${formatDate(action.dueDate)} · ${action.owner ?? "No owner"}`, status: action.status }))} />
      </div>
    </>
  );
}

function TeamActionTrackerView({
  schedule,
  tracker,
  dateWindow,
  onExportPdf,
}: {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  onExportPdf: (items: TeamWorkItem[]) => void;
}) {
  const [statusFilters, setStatusFilters] = useState<TeamStatusFilter[]>(["open", "due-soon", "overdue", "blocked"]);
  const [owner, setOwner] = useState("all");
  const [source, setSource] = useState("all");
  const [stream, setStream] = useState("all");
  const [search, setSearch] = useState("");
  const allItems = combineTeamWorkItems(schedule, tracker);
  const owners = uniqueSorted(allItems.flatMap((item) => item.owner?.split("/") ?? []).map((value) => value.trim()).filter(Boolean));
  const streams = uniqueSorted(allItems.map((item) => item.stream));
  const sources = uniqueSorted(allItems.map((item) => item.source));
  const filtered = allItems
    .filter((item) => {
      const group = actionStatusGroup(item);
      const dateForWindow = group === "completed" ? item.completionDate ?? item.dueDate ?? item.meetingDate : item.dueDate ?? item.meetingDate;
      if (!teamStatusMatches(group, statusFilters)) return false;
      if (owner !== "all" && !normaliseText(item.owner).includes(normaliseText(owner))) return false;
      if (source !== "all" && item.source !== source) return false;
      if (stream !== "all" && item.stream !== stream) return false;
      if (search && !normaliseText(`${item.title} ${item.description} ${item.latestUpdate} ${item.owner} ${item.stream}`).includes(normaliseText(search))) return false;
      if (!dateWithin(dateForWindow, dateWindow)) return false;
      return true;
    })
    .sort((a, b) => bySoonest(a.dueDate, b.dueDate) || a.title.localeCompare(b.title));
  const counts = {
    open: allItems.filter((item) => actionStatusGroup(item) !== "completed").length,
    dueSoon: allItems.filter((item) => actionStatusGroup(item) === "due-soon").length,
    overdue: allItems.filter((item) => actionStatusGroup(item) === "overdue").length,
    blocked: allItems.filter((item) => actionStatusGroup(item) === "blocked").length,
    completed: allItems.filter((item) => actionStatusGroup(item) === "completed").length,
  };
  const toggleStatusFilter = (key: TeamStatusFilter) => {
    setStatusFilters((current) => current.includes(key) ? current.filter((item) => item !== key) : [...current, key]);
  };

  return (
    <>
      <div id="team-actions-export" className="team-actions-view">
        <PageIntro title="Team Action Tracker" tracker={tracker}>A combined operational view of meeting actions and assigned Project plan tasks or milestones.</PageIntro>
        <StatGrid cards={[
          ["Open / needs doing", counts.open.toString()],
          ["Due soon", counts.dueSoon.toString(), counts.dueSoon ? "warn" : ""],
          ["Overdue", counts.overdue.toString(), counts.overdue ? "warn" : ""],
          ["Completed", counts.completed.toString()],
        ]} />

        <section className="team-action-controls">
          <div className="tabs" aria-label="Action status filters">
            {teamStatusFilters.map(({ key, label }) => (
              <button
                type="button"
                className={statusFilters.includes(key) ? "active" : ""}
                aria-pressed={statusFilters.includes(key)}
                onClick={() => toggleStatusFilter(key)}
                key={key}
              >
                {label}
              </button>
            ))}
            <button type="button" className={!statusFilters.length ? "active" : ""} aria-pressed={!statusFilters.length} onClick={() => setStatusFilters([])}>All</button>
          </div>
          <div className="team-filter-grid">
            <label>
              Owner
              <select value={owner} onChange={(event) => setOwner(event.target.value)}>
                <option value="all">All owners</option>
                {owners.map((value) => <option key={value} value={value}>{value}</option>)}
              </select>
            </label>
            <label>
              Source
              <select value={source} onChange={(event) => setSource(event.target.value)}>
                <option value="all">All sources</option>
                {sources.map((value) => <option key={value} value={value}>{value}</option>)}
              </select>
            </label>
            <label>
              Workstream
              <select value={stream} onChange={(event) => setStream(event.target.value)}>
                <option value="all">All workstreams</option>
                {streams.map((value) => <option key={value} value={value}>{value}</option>)}
              </select>
            </label>
            <label>
              Search
              <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Action, task, owner..." />
            </label>
          </div>
        </section>

        <section className="team-action-list">
          {filtered.map((item) => {
            const group = actionStatusGroup(item);
            return (
              <details className={`team-action-card action-${group}`} key={item.id}>
                <summary>
                  <span className="source-chip">{item.source}</span>
                  <strong>{item.title}</strong>
                  <span>{item.owner ?? "No owner"}</span>
                  <span className="team-action-date"><small>Logged</small>{formatDate(item.loggedDate ?? item.meetingDate)}</span>
                  <span className="team-action-date"><small>Due</small>{formatDate(item.dueDate)}</span>
                  <em>{statusLabel(item)}</em>
                </summary>
                <div className="team-action-detail">
                  <p>{item.description || "No detailed description held."}</p>
                  {item.latestUpdate ? <p><strong>Latest update:</strong> {item.latestUpdate}</p> : null}
                  <dl>
                    <div><dt>Workstream</dt><dd>{item.stream ?? "Not set"}</dd></div>
                    <div><dt>Priority</dt><dd>{item.priority ?? "Not set"}</dd></div>
                    <div><dt>Logged date</dt><dd>{formatDate(item.loggedDate ?? item.meetingDate)}</dd></div>
                    <div><dt>Due date</dt><dd>{formatDate(item.dueDate)}</dd></div>
                    <div><dt>Completion date</dt><dd>{item.completionDate ? formatDate(item.completionDate) : "Not held"}</dd></div>
                    <div><dt>Links</dt><dd>{item.links || "None"}</dd></div>
                  </dl>
                </div>
              </details>
            );
          })}
          {!filtered.length ? <article className="empty-panel"><h2>No actions found</h2><p>Adjust the status, owner, source, workstream or date window filters.</p></article> : null}
        </section>
      </div>
      <div className="snapshot-actions">
        <button className="download-action" type="button" onClick={() => onExportPdf(filtered)}>
          <Download size={15} />
          Download A4 Actions PDF
        </button>
        <button className="download-action" type="button" onClick={() => downloadTeamActionsCsv(filtered, schedule)}>
          <Download size={15} />
          Download CSV
        </button>
      </div>
    </>
  );
}

function DependencyView({ schedule, tracker }: { schedule: ProgrammeSchedule; tracker?: TrackerData }) {
  const dependencies = schedule.items.filter((item) => item.dependencyLevel || item.externalDependency || item.predecessors.length || item.successors.length);
  const blockers = [...openRisks(tracker), ...openIssues(tracker)].filter((item) => item.dashboardFlag || isRedOrAmber("rag" in item ? item.rag : undefined));
  return (
    <>
      <PageIntro title="Dependency View" tracker={tracker}>A first dependency readout using Project predecessor/successor links and tracker blockers.</PageIntro>
      <div className="report-grid two">
        <ItemList title="Executive and Programme Dependencies" items={dependencies.slice(0, 10).map((item) => ({ id: item.uid, title: item.name, meta: `${item.dependencyLevel ?? "Linked dependency"} · ${item.stream ?? "No stream"}`, status: item.status }))} />
        <ItemList title="Blocking Items from Tracker" items={blockers.map((item) => ({ id: item.id, title: item.title, meta: item.latestUpdate, status: "rag" in item ? item.rag ?? item.status : item.status }))} />
      </div>
    </>
  );
}

function WorkstreamViews({ schedule, tracker }: { schedule: ProgrammeSchedule; tracker?: TrackerData }) {
  const streams = uniqueSorted([
    ...schedule.items.map((item) => item.stream),
    ...(tracker?.risks.map((risk) => risk.stream) ?? []),
    ...(tracker?.issues.map((issue) => issue.stream) ?? []),
    ...(tracker?.actions.map((action) => action.stream) ?? []),
  ]);
  return (
    <>
      <PageIntro title="Workstream Views" tracker={tracker}>A stream-by-stream reporting index for leads, using the same underlying project and tracker data.</PageIntro>
      <section className="workstream-grid">
        {streams.map((stream) => {
          const next = schedule.items.filter((item) => item.stream === stream && item.status !== "complete").sort((a, b) => bySoonest(a.finishDate, b.finishDate))[0];
          const risks = openRisks(tracker).filter((risk) => risk.stream === stream).length;
          const actions = openActions(tracker).filter((action) => action.stream === stream).length;
          return (
            <article className="report-card" key={stream}>
              <h3>{stream}</h3>
              <p><strong>Next milestone:</strong> {next ? `${formatDate(next.finishDate)} - ${next.name}` : "None found"}</p>
              <p><strong>Open risks:</strong> {risks}</p>
              <p><strong>Open actions:</strong> {actions}</p>
            </article>
          );
        })}
      </section>
    </>
  );
}

function PartnerView({ schedule, tracker }: { schedule: ProgrammeSchedule; tracker?: TrackerData }) {
  const partnerItems = schedule.items.filter((item) => /partner|working group|adopter|industry|pilot/i.test(`${item.name} ${item.stream}`)).slice(0, 10);
  return (
    <>
      <PageIntro title="Partner View" tracker={tracker}>A restricted-audience view scaffold for partner timelines, asks and working group dates.</PageIntro>
      <div className="report-grid two">
        <ItemList title="Partner Timeline" items={partnerItems.map((item) => ({ id: item.uid, title: item.name, meta: `${formatDate(item.finishDate)} · ${item.stream ?? "No stream"}`, status: item.status }))} />
        <ItemList title="Partner Actions" items={openActions(tracker).filter((action) => /partner|working group|adopter|industry|pilot/i.test(`${action.title} ${action.description} ${action.stream}`)).map((action) => ({ id: action.id, title: action.title, meta: `${formatDate(action.dueDate)} · ${action.owner ?? "No owner"}`, status: action.status }))} />
      </div>
    </>
  );
}

function ReleasePlaceholder({ title }: { title: string }) {
  return (
    <section className="empty-panel">
      <h2>{title}</h2>
      <p>The Release Plan data source has intentionally not been inferred from the current XML or tracker. This page is reserved for the release plan file once it is available.</p>
      <p>The build now has a separate release-plan slot, so versions, scope, acceptance criteria and readiness can be added cleanly without reworking the current reporting pages.</p>
    </section>
  );
}

function DownloadsHub({
  schedule,
  tracker,
  dateWindow,
  contextItemUids,
  removedItemUids,
  laneOrderUids,
  onExportPdf,
  onExportPosterPdf,
  onExportJson,
  onExportSnapshotPdf,
  onExportSnapshotImage,
  onExportSnapshotPosterPdf,
  onExportSnapshotHtml,
}: {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  contextItemUids: string[];
  removedItemUids: string[];
  laneOrderUids: string[];
  onExportPdf: () => void;
  onExportPosterPdf: () => void;
  onExportJson: () => void;
  onExportSnapshotPdf: () => void;
  onExportSnapshotImage: () => void;
  onExportSnapshotPosterPdf: () => void;
  onExportSnapshotHtml: () => void;
}) {
  const downloads = [
    {
      title: "Programme roadmap PDF",
      meta: "Exports the current roadmap workspace view using the selected filters and date window.",
      action: "Download PDF",
      onClick: onExportPdf,
    },
    {
      title: "Programme poster PDF",
      meta: "Exports the visual poster timeline for senior stakeholder sharing.",
      action: "Download Poster",
      onClick: onExportPosterPdf,
    },
    {
      title: "Normalised schedule JSON",
      meta: `${schedule.items.length} plan items available from the imported XML.`,
      action: "Download JSON",
      onClick: onExportJson,
    },
    {
      title: "Executive view PDF",
      meta: "Exports the detailed A4 table pack for the Executive delivery roadmap.",
      action: "Download Table PDF",
      onClick: onExportSnapshotPdf,
    },
    {
      title: "Executive roadmap poster PDF",
      meta: "Exports a designed print-safe executive roadmap poster without slicing the web page.",
      action: "Download Poster PDF",
      onClick: onExportSnapshotPosterPdf,
    },
    {
      title: "Executive roadmap image",
      meta: "Exports the visible Executive View as a high-resolution PNG for Teams, email or slides.",
      action: "Download PNG",
      onClick: onExportSnapshotImage,
    },
    {
      title: "Executive standalone HTML",
      meta: "Exports a self-contained browser file with the executive roadmap and print styling.",
      action: "Download HTML",
      onClick: onExportSnapshotHtml,
    },
    {
      title: "Board report PDF",
      meta: "Planned export for board reporting packs.",
      action: "Coming soon",
      disabled: true,
    },
    {
      title: "Risk and issue report",
      meta: tracker ? "Tracker data imported and ready for a future export." : "Import tracker first.",
      action: "Coming soon",
      disabled: true,
    },
    {
      title: "Gantt extract",
      meta: "Planned extract for schedule analysis.",
      action: "Coming soon",
      disabled: true,
    },
    {
      title: "Partner roadmap",
      meta: "Planned export for partner-facing roadmap views.",
      action: "Coming soon",
      disabled: true,
    },
  ];

  return (
    <>
      <PageIntro title="Downloads" tracker={tracker}>A central location for current exports and future audience-specific report packs.</PageIntro>
      <section className="download-grid">
        {downloads.map((item) => (
          <article className="report-card download-card" key={item.title}>
            <div>
              <h3>{item.title}</h3>
              <p>{item.meta}</p>
            </div>
            <button className="download-action" type="button" onClick={item.onClick} disabled={item.disabled}>
              <Download size={15} />
              {item.action}
            </button>
          </article>
        ))}
      </section>
      <div className="snapshot-export-mount" aria-hidden="true">
        <ExecutiveSnapshotView schedule={schedule} tracker={tracker} dateWindow={dateWindow} contextItemUids={contextItemUids} removedItemUids={removedItemUids} laneOrderUids={laneOrderUids} />
      </div>
    </>
  );
}

function ReportingContent({
  page,
  schedule,
  tracker,
  dateWindow,
  selected,
  setSelected,
  onExportPdf,
  onExportPosterPdf,
  onExportJson,
  onExportSnapshotPdf,
  onExportSnapshotImage,
  onExportSnapshotPosterPdf,
  onExportSnapshotHtml,
  executiveContextUids,
  executiveRemovedUids,
  executiveLaneOrderUids,
  onToggleExecutiveContextItem,
  onRemoveExecutiveRoadmapItem,
  onRestoreExecutiveRoadmapItem,
  onReorderExecutiveLanes,
  weeklyStatusCuration,
  onUpdateWeeklyStatusCuration,
  onExportWeeklyStatusPdf,
  onExportTeamActionsPdf,
}: {
  page: AppPage;
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  selected?: ProgrammeItem;
  setSelected: (item: ProgrammeItem) => void;
  onExportPdf: () => void;
  onExportPosterPdf: () => void;
  onExportJson: () => void;
  onExportSnapshotPdf: () => void;
  onExportSnapshotImage: () => void;
  onExportSnapshotPosterPdf: () => void;
  onExportSnapshotHtml: () => void;
  executiveContextUids: string[];
  executiveRemovedUids: string[];
  executiveLaneOrderUids: string[];
  onToggleExecutiveContextItem: (uid: string) => void;
  onRemoveExecutiveRoadmapItem: (uid: string) => void;
  onRestoreExecutiveRoadmapItem: (uid: string) => void;
  onReorderExecutiveLanes: (orderedUids: string[]) => void;
  weeklyStatusCuration: WeeklyStatusCuration;
  onUpdateWeeklyStatusCuration: (updater: (current: WeeklyStatusCuration) => WeeklyStatusCuration) => void;
  onExportWeeklyStatusPdf: () => void;
  onExportTeamActionsPdf: (items: TeamWorkItem[]) => void;
}) {
  if (page === "home") return <HomeDashboard schedule={schedule} tracker={tracker} dateWindow={dateWindow} />;
  if (page === "ceo") return (
    <ExecutiveSnapshotView
      schedule={schedule}
      tracker={tracker}
      dateWindow={dateWindow}
      contextItemUids={executiveContextUids}
      removedItemUids={executiveRemovedUids}
      laneOrderUids={executiveLaneOrderUids}
      onToggleContextItem={onToggleExecutiveContextItem}
      onRemoveRoadmapItem={onRemoveExecutiveRoadmapItem}
      onRestoreRoadmapItem={onRestoreExecutiveRoadmapItem}
      onReorderLanes={onReorderExecutiveLanes}
      onExportSnapshotPdf={onExportSnapshotPdf}
      onExportSnapshotImage={onExportSnapshotImage}
      onExportSnapshotPosterPdf={onExportSnapshotPosterPdf}
      onExportSnapshotHtml={onExportSnapshotHtml}
    />
  );
  if (page === "weekly-status") return (
    <WeeklyExecutiveStatusView
      schedule={schedule}
      tracker={tracker}
      dateWindow={dateWindow}
      curation={weeklyStatusCuration}
      onUpdateCuration={onUpdateWeeklyStatusCuration}
      onExportPdf={onExportWeeklyStatusPdf}
    />
  );
  if (page === "team-actions") return <TeamActionTrackerView schedule={schedule} tracker={tracker} dateWindow={dateWindow} onExportPdf={onExportTeamActionsPdf} />;
  if (page === "board") return <BoardReportView schedule={schedule} tracker={tracker} dateWindow={dateWindow} />;
  if (page === "reporting-roadmap") return <ReportingRoadmapView schedule={schedule} dateWindow={dateWindow} selected={selected} onSelect={setSelected} />;
  if (page === "gantt") return <GanttView schedule={schedule} dateWindow={dateWindow} selected={selected} onSelect={setSelected} />;
  if (page === "risks") return <RisksIssuesView tracker={tracker} />;
  if (page === "actions") return <ActionsDecisionsView schedule={schedule} tracker={tracker} dateWindow={dateWindow} />;
  if (page === "dependencies") return <DependencyView schedule={schedule} tracker={tracker} />;
  if (page === "workstreams") return <WorkstreamViews schedule={schedule} tracker={tracker} />;
  if (page === "partner") return <PartnerView schedule={schedule} tracker={tracker} />;
  if (page === "downloads") return (
    <DownloadsHub
      schedule={schedule}
      tracker={tracker}
      dateWindow={dateWindow}
      contextItemUids={executiveContextUids}
      removedItemUids={executiveRemovedUids}
      laneOrderUids={executiveLaneOrderUids}
      onExportPdf={onExportPdf}
      onExportPosterPdf={onExportPosterPdf}
      onExportJson={onExportJson}
      onExportSnapshotPdf={onExportSnapshotPdf}
      onExportSnapshotImage={onExportSnapshotImage}
      onExportSnapshotPosterPdf={onExportSnapshotPosterPdf}
      onExportSnapshotHtml={onExportSnapshotHtml}
    />
  );
  if (page === "release-roadmap") return <ReleasePlaceholder title="Release Roadmap" />;
  if (page === "version-scope") return <ReleasePlaceholder title="Version Scope" />;
  if (page === "release-readiness") return <ReleasePlaceholder title="Release Readiness" />;
  return null;
}

function DetailDrawer({ item, onClose }: { item?: ProgrammeItem; onClose: () => void }) {
  if (!item) return null;
  const fields = [
    ["Stream", item.stream],
    ["WBS / Outline", item.wbs ?? item.outlineNumber],
    ["Type", item.itemType],
    ["Start", formatDate(item.startDate)],
    ["Finish", formatDate(item.finishDate)],
    ["Baseline finish", formatDate(item.baselineFinish)],
    ["Delay", item.delayDays ? `${item.delayDays} calendar days` : "No delay"],
    ["Percent complete", `${item.percentComplete ?? 0}%`],
    ["Critical", item.isCritical ? "Yes" : "No"],
    ["Milestone type", item.milestoneType],
    ["Approval body", item.approvalBody],
    ["Version", item.version],
    ["Visibility", item.visibility],
    ["Roadmap view", item.roadmapView],
    ["Resources", item.resourceNames?.join(", ")],
    ["Predecessors", item.predecessors.map((dependency) => dependency.predecessorUid).filter(Boolean).join(", ")],
    ["Successors", item.successors.map((dependency) => dependency.successorUid).filter(Boolean).join(", ")],
  ];
  return (
    <aside className="drawer">
      <button className="icon-button close" type="button" onClick={onClose} title="Close details"><X size={18} /></button>
      <span className={`status-pill ${item.status}`}>{item.status}</span>
      <h2>{item.name}</h2>
      <div className="drawer-grid">
        {fields.map(([label, value]) => (
          <React.Fragment key={label}>
            <span>{label}</span>
            <strong>{value || "Not set"}</strong>
          </React.Fragment>
        ))}
      </div>
    </aside>
  );
}

function App() {
  const [schedule, setSchedule] = useState<ProgrammeSchedule>(makeSampleSchedule());
  const [tracker, setTracker] = useState<TrackerData | undefined>();
  const [filters, setFilters] = useState<ProgrammeFilters>(initialFilters);
  const [view, setView] = useState<ProgrammeView>("roadmap");
  const [page, setPage] = useState<AppPage>("workspace");
  const [selected, setSelected] = useState<ProgrammeItem | undefined>();
  const [error, setError] = useState<string | undefined>();
  const [baselineNumber, setBaselineNumber] = useState(3);
  const [sourceXml, setSourceXml] = useState<{ xml: string; fileName: string } | undefined>();
  const [executiveContextUids, setExecutiveContextUids] = useState<string[]>([]);
  const [executiveRemovedUids, setExecutiveRemovedUids] = useState<string[]>([]);
  const [executiveLaneOrderUids, setExecutiveLaneOrderUids] = useState<string[]>([]);
  const [weeklyStatusCuration, setWeeklyStatusCuration] = useState<WeeklyStatusCuration>({});
  const [theme, setTheme] = useState<"light" | "dark">(() => {
    const stored = window.localStorage.getItem("roadmap-theme");
    if (stored === "light" || stored === "dark") return stored;
    return window.matchMedia?.("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  });

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    window.localStorage.setItem("roadmap-theme", theme);
  }, [theme]);

  useEffect(() => {
    setFilters((current) => {
      if (current.dateStart || current.dateEnd) return current;
      return {
        ...current,
        dateStart: inputDate(schedule.statusDate ?? schedule.startDate),
        dateEnd: inputDate(schedule.finishDate),
      };
    });
  }, [schedule]);

  const dateWindow = useMemo(() => resolveDateWindow(filters, schedule), [filters, schedule]);
  const visibleItems = useMemo(() => applyFilters(applyView(schedule.items, view), filters, schedule), [schedule, view, filters]);

  useEffect(() => {
    const validUids = new Set(schedule.items.map((item) => item.uid));
    setExecutiveContextUids((current) => current.filter((uid) => validUids.has(uid)));
    setExecutiveRemovedUids((current) => current.filter((uid) => validUids.has(uid)));
    setExecutiveLaneOrderUids((current) => current.filter((uid) => validUids.has(uid)));
  }, [schedule.items]);

  useEffect(() => {
    if (!sourceXml) return;
    try {
      setSchedule(parseMicrosoftProjectXml(sourceXml.xml, sourceXml.fileName, baselineNumber));
    } catch {
      return;
    }
  }, [baselineNumber, sourceXml]);

  async function importFile(file: File) {
    setError(undefined);
    try {
      const xml = await file.text();
      const parsed = parseMicrosoftProjectXml(xml, file.name, baselineNumber);
      setSourceXml({ xml, fileName: file.name });
      setSchedule(parsed);
      setSelected(undefined);
      setExecutiveContextUids([]);
      setExecutiveRemovedUids([]);
      setExecutiveLaneOrderUids([]);
      setWeeklyStatusCuration({});
    } catch (err) {
      setError(err instanceof Error ? err.message : "The file could not be imported.");
    }
  }

  async function importTracker(file: File) {
    setError(undefined);
    try {
      const parsed = await parseMeetingTracker(file);
      setTracker(parsed);
      setWeeklyStatusCuration({});
    } catch (err) {
      setError(err instanceof Error ? err.message : "The tracker workbook could not be imported.");
    }
  }

  function exportJson() {
    const blob = new Blob([JSON.stringify(schedule, null, 2)], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${schedule.title.replace(/[^a-z0-9]+/gi, "-").toLowerCase()}-normalised-schedule.json`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  function toggleExecutiveContextItem(uid: string) {
    setExecutiveContextUids((current) => current.includes(uid) ? current.filter((item) => item !== uid) : [...current, uid]);
    setExecutiveRemovedUids((current) => current.filter((item) => item !== uid));
  }

  function removeExecutiveRoadmapItem(uid: string) {
    setExecutiveContextUids((current) => current.filter((item) => item !== uid));
    setExecutiveRemovedUids((current) => current.includes(uid) ? current : [...current, uid]);
  }

  function restoreExecutiveRoadmapItem(uid: string) {
    setExecutiveRemovedUids((current) => current.filter((item) => item !== uid));
  }

  function reorderExecutiveLanes(orderedUids: string[]) {
    setExecutiveLaneOrderUids(orderedUids);
  }

  async function exportPdf() {
    setError(undefined);
    try {
      await exportProgrammePdf({
        schedule,
        items: visibleItems,
        viewLabel: viewLabels[view],
        filters,
        dateWindowLabel: dateWindow.label,
        dateWindowStart: dateWindow.start?.toISOString(),
        dateWindowEnd: dateWindow.end?.toISOString(),
        baselineNumber,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "The PDF report could not be generated.");
    }
  }

  async function exportPosterPdf() {
    setError(undefined);
    try {
      await exportProgrammePdf({
        schedule,
        items: visibleItems,
        viewLabel: viewLabels[view],
        filters,
        dateWindowLabel: dateWindow.label,
        dateWindowStart: dateWindow.start?.toISOString(),
        dateWindowEnd: dateWindow.end?.toISOString(),
        baselineNumber,
        output: "poster",
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "The poster PDF could not be generated.");
    }
  }

  async function exportExecutiveSnapshotPdf() {
    setError(undefined);
    try {
      await exportExecutiveRoadmapPdf({ schedule, tracker, dateWindow, contextItemUids: executiveContextUids, removedItemUids: executiveRemovedUids, laneOrderUids: executiveLaneOrderUids });
    } catch (err) {
      setError(err instanceof Error ? err.message : "The Executive View PDF could not be generated.");
    }
  }

  async function exportExecutiveSnapshotImage() {
    setError(undefined);
    try {
      const element = document.getElementById("executive-snapshot-export");
      if (!element) throw new Error("Open the Executive View page or Downloads page before exporting the executive roadmap image.");
      await exportExecutiveRoadmapImage(element, schedule);
    } catch (err) {
      setError(err instanceof Error ? err.message : "The Executive View image could not be generated.");
    }
  }

  async function exportExecutiveSnapshotPosterPdf() {
    setError(undefined);
    try {
      await exportExecutiveRoadmapPosterPdf(schedule, tracker, dateWindow, executiveContextUids, executiveRemovedUids, executiveLaneOrderUids);
    } catch (err) {
      setError(err instanceof Error ? err.message : "The Executive roadmap poster PDF could not be generated.");
    }
  }

  function exportExecutiveSnapshotHtml() {
    setError(undefined);
    try {
      exportExecutiveRoadmapHtml(schedule, tracker, dateWindow, executiveContextUids, executiveRemovedUids, executiveLaneOrderUids);
    } catch (err) {
      setError(err instanceof Error ? err.message : "The Executive roadmap HTML file could not be generated.");
    }
  }

  async function exportWeeklyStatusPdf() {
    setError(undefined);
    try {
      await exportWeeklyStatusA4Pdf({ schedule, tracker, dateWindow, curation: weeklyStatusCuration });
    } catch (err) {
      setError(err instanceof Error ? err.message : "The Weekly Executive Status PDF could not be generated.");
    }
  }

  async function exportTeamActionsPdf(items: TeamWorkItem[]) {
    setError(undefined);
    try {
      await exportTeamActionsA4Pdf({
        schedule,
        dateWindow,
        items: items.map((item) => ({
          ...item,
          displayStatus: statusLabel(item),
        })),
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "The Team Action Tracker PDF could not be generated.");
    }
  }

  return (
    <main>
      <header className="app-header">
        <div>
          <span className="eyebrow"><ShieldCheck size={16} /> Local browser-only schedule intelligence</span>
          <h1>Interactive Programme Roadmap / Integrated Master Schedule Generator</h1>
          <p>{schedule.title}{schedule.sourceFileName ? ` from ${schedule.sourceFileName}` : ""}</p>
        </div>
        <div className="header-actions">
          <button className="theme-button" type="button" onClick={() => setTheme((current) => current === "dark" ? "light" : "dark")}>
            {theme === "dark" ? <Sun size={18} /> : <Moon size={18} />}
            {theme === "dark" ? "Light mode" : "Dark mode"}
          </button>
          <label className="upload-button">
            <FileUp size={18} />
            Import Project XML
            <input type="file" accept=".xml,text/xml,application/xml" onChange={(event) => event.target.files?.[0] && importFile(event.target.files[0])} />
          </label>
          <label className="upload-button secondary">
            <FileSpreadsheet size={18} />
            Import Tracker XLSX
            <input type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" onChange={(event) => event.target.files?.[0] && importTracker(event.target.files[0])} />
          </label>
        </div>
      </header>

      {error ? <div className="error">{error}</div> : null}
      <SummaryBar schedule={schedule} tracker={tracker} />
      <nav className="app-nav" aria-label="Application sections">
        {appPages.map((item) => (
          <button type="button" className={`${page === item.key ? "active" : ""} ${item.group}`} onClick={() => setPage(item.key)} key={item.key}>
            {item.key === "workspace" ? <Layers size={15} /> : item.key === "gantt" ? <GitBranch size={15} /> : item.group === "future" ? <CalendarDays size={15} /> : item.key === "workstreams" || item.key === "partner" ? <Users size={15} /> : item.key === "actions" ? <ClipboardCheck size={15} /> : <LayoutDashboard size={15} />}
            {item.label}
          </button>
        ))}
      </nav>

      {page === "workspace" ? (
        <div className="workspace">
          <FilterPanel schedule={schedule} filters={filters} setFilters={setFilters} />
          <section className="content">
            <div className="view-bar">
              <div className="tabs">
                {(Object.keys(viewLabels) as ProgrammeView[]).map((key) => (
                  <button type="button" className={view === key ? "active" : ""} onClick={() => setView(key)} key={key}>
                    {key === "roadmap" ? <Layers size={15} /> : key === "milestones" ? <Milestone size={15} /> : key === "governance" ? <ShieldCheck size={15} /> : key === "schedule" ? <GitBranch size={15} /> : <CalendarDays size={15} />}
                    {viewLabels[key]}
                  </button>
                ))}
              </div>
              <div className="controls">
                <label>
                  <SlidersHorizontal size={15} />
                  Baseline
                  <select value={baselineNumber} onChange={(event) => setBaselineNumber(Number(event.target.value))}>
                    {[0, 1, 2, 3].map((baseline) => <option key={baseline} value={baseline}>Baseline {baseline}</option>)}
                  </select>
                </label>
                <button type="button" onClick={exportPdf}><Download size={15} /> PDF</button>
                <button type="button" onClick={exportPosterPdf}><Download size={15} /> Poster PDF</button>
                <button type="button" onClick={exportJson}><Download size={15} /> JSON</button>
              </div>
            </div>
            <Timeline schedule={schedule} items={visibleItems} selected={selected} onSelect={setSelected} dateWindow={dateWindow} />
            <InsightsPanel schedule={schedule} />
          </section>
        </div>
      ) : (
        <section className="reporting-shell">
          <ReportingPeriodControl filters={filters} setFilters={setFilters} />
          <ReportingContent
            page={page}
            schedule={schedule}
            tracker={tracker}
            dateWindow={dateWindow}
            selected={selected}
            setSelected={setSelected}
            onExportPdf={exportPdf}
            onExportPosterPdf={exportPosterPdf}
            onExportJson={exportJson}
            onExportSnapshotPdf={exportExecutiveSnapshotPdf}
            onExportSnapshotImage={exportExecutiveSnapshotImage}
            onExportSnapshotPosterPdf={exportExecutiveSnapshotPosterPdf}
            onExportSnapshotHtml={exportExecutiveSnapshotHtml}
            executiveContextUids={executiveContextUids}
            executiveRemovedUids={executiveRemovedUids}
            executiveLaneOrderUids={executiveLaneOrderUids}
            onToggleExecutiveContextItem={toggleExecutiveContextItem}
            onRemoveExecutiveRoadmapItem={removeExecutiveRoadmapItem}
            onRestoreExecutiveRoadmapItem={restoreExecutiveRoadmapItem}
            onReorderExecutiveLanes={reorderExecutiveLanes}
            weeklyStatusCuration={weeklyStatusCuration}
            onUpdateWeeklyStatusCuration={setWeeklyStatusCuration}
            onExportWeeklyStatusPdf={exportWeeklyStatusPdf}
            onExportTeamActionsPdf={exportTeamActionsPdf}
          />
        </section>
      )}
      <DetailDrawer item={selected} onClose={() => setSelected(undefined)} />
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
