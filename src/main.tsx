import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  AlertTriangle,
  ArrowRight,
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
import { exportElementPdf } from "./lib/exportElementPdf";
import { parseMicrosoftProjectXml } from "./lib/parseMicrosoftProjectXml";
import { parseMeetingTracker } from "./lib/parseMeetingTracker";
import { clamp, formatDate, parseDate, uniqueSorted } from "./lib/dateUtils";
import type { ProgrammeFilters, ProgrammeItem, ProgrammeSchedule, ProgrammeView } from "./types/programme";
import type { TrackerAction, TrackerData, TrackerDecision, TrackerIssue, TrackerRisk } from "./types/reporting";
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
  | "board"
  | "reporting-roadmap"
  | "risks"
  | "actions"
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
  { key: "board", label: "Board Report", group: "reporting" },
  { key: "reporting-roadmap", label: "Reporting Roadmap", group: "reporting" },
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

function SummaryBar({ schedule }: { schedule: ProgrammeSchedule }) {
  const totalMilestones = schedule.items.filter((item) => item.isMilestone).length;
  const roadmapMilestones = schedule.items.filter((item) => item.roadmapMilestone).length;
  const critical = schedule.items.filter((item) => item.isCritical).length;
  const delayed = schedule.items.filter((item) => item.delayDays && item.delayDays > 0).length;
  const cards = [
    ["Programme start", formatDate(schedule.startDate)],
    ["Programme finish", formatDate(schedule.finishDate)],
    ["Status date", formatDate(schedule.statusDate)],
    ["Total tasks", schedule.items.length.toString()],
    ["Milestones", totalMilestones.toString()],
    ["Roadmap", roadmapMilestones.toString()],
    ["Critical", critical.toString()],
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

function splitDigest(value?: string, limit = 3): string[] {
  const text = value?.trim();
  if (!text) return [];
  const lineParts = text.split(/\n+/).map((part) => part.trim()).filter(Boolean);
  const parts = lineParts.length > 1 ? lineParts : text.split(/(?<=[.!?])\s+/).map((part) => part.trim()).filter(Boolean);
  return parts.slice(0, limit);
}

function formatDateOrText(value?: string, fallback = "Not set"): string {
  if (!value) return fallback;
  return parseDate(value) ? formatDate(value) : value;
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
  return tracker?.weeklySummaries
    .slice()
    .sort((a, b) => (parseDate(b.weekEnding)?.getTime() ?? 0) - (parseDate(a.weekEnding)?.getTime() ?? 0))[0];
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
    .filter((item) => item.executiveMilestone || normaliseText(item.milestoneLevel).includes("executive"))
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate));
  if (executive.length) return executive;
  return programmeMilestones(schedule).filter((item) => itemImportance(item) >= 4).slice(0, 5);
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

function importantExecutiveDependency(item: ProgrammeItem): boolean {
  return executiveDependencyScore(item) > 0;
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

function collectPredecessorDependencies(outcome: ProgrammeItem, byUid: Map<string, ProgrammeItem>): ProgrammeItem[] {
  return collectPredecessorChain(outcome, byUid).filter(importantExecutiveDependency);
}

function executiveDependencyPaths(schedule: ProgrammeSchedule): ExecutivePath[] {
  const outcomes = executiveMilestoneItems(schedule).slice(0, 7);
  const byUid = new Map(schedule.items.map((item) => [item.uid, item]));
  return outcomes.map((outcome) => {
    const predecessors = collectPredecessorDependencies(outcome, byUid);
    const combined = new Map<string, ProgrammeItem>();
    predecessors.forEach((item) => combined.set(item.uid, item));
    const priorityDependencies = [...combined.values()]
      .sort((a, b) => executiveDependencyScore(b) - executiveDependencyScore(a) || bySoonest(a.finishDate, b.finishDate))
      .slice(0, 5);
    const dependencies = priorityDependencies.sort((a, b) => bySoonest(a.finishDate, b.finishDate) || executiveDependencyScore(b) - executiveDependencyScore(a));
    return { outcome, dependencies };
  });
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
  onExportSnapshotPdf,
}: {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  onExportSnapshotPdf?: () => void;
}) {
  const paths = executiveDependencyPaths(schedule);
  const weekly = latestWeeklySummary(tracker);
  const [selectedUid, setSelectedUid] = useState(paths[0]?.outcome.uid ?? "");
  const selectedPath = paths.find((path) => path.outcome.uid === selectedUid) ?? paths[0];
  const outcomes = paths.map((path) => path.outcome);
  const programmeTone = toneClass(weekly?.overallRag);
  const confidenceTone = recoveryConfidence(weekly, outcomes);
  const reportingDate = new Date().toISOString();
  const keyEnablerWindow = { ...dateWindow, start: dateWindow.start ?? parseDate(reportingDate) };
  const byUid = new Map(schedule.items.map((item) => [item.uid, item]));
  const programmeStatusText = weekly?.overallRag ? weekly.overallRag.toUpperCase() : "Not set";
  const programmeReason = normaliseText(weekly?.overallRag).includes("red")
    ? "Original July 2026 programme commitment missed"
    : weekly?.ragRationale ?? "Import the latest meeting tracker to populate the current programme position.";
  const decisions = sortFlaggedFirst(openDecisions(tracker)).slice(0, 4);
  const keyEnablers = [...new Map(paths.flatMap((path) => collectPredecessorChain(path.outcome, byUid)).map((item) => [item.uid, item])).values()]
    .filter((item) => !item.executiveMilestone && !item.isSummary && item.isActive && overlapsDateWindow(item, keyEnablerWindow) && executiveEnablerScore(item) > 0)
    .sort((a, b) => bySoonest(a.finishDate, b.finishDate) || executiveEnablerScore(b) - executiveEnablerScore(a))
    .slice(0, 5);
  const watchlistItems = sortFlaggedFirst([
    ...openRisks(tracker)
      .filter((risk) => risk.dashboardFlag || isRedOrAmber(risk.rag))
      .map((risk) => ({ id: `risk-${risk.id}`, title: risk.title, meta: risk.rag ?? risk.status ?? "Risk", dashboardFlag: risk.dashboardFlag })),
    ...openIssues(tracker)
      .filter((issue) => issue.dashboardFlag || isRedOrAmber(issue.rag) || isRedOrAmber(issue.priority))
      .map((issue) => ({ id: `issue-${issue.id}`, title: issue.title, meta: issue.rag ?? issue.priority ?? issue.status ?? "Issue", dashboardFlag: issue.dashboardFlag })),
  ]).slice(0, 5);
  const attentionItems = [
    weekly?.askSteerNeeded ? { id: "ask", title: weekly.askSteerNeeded, meta: "Current ask" } : undefined,
    weekly?.mainBlocker ? { id: "blocker", title: weekly.mainBlocker, meta: "Main blocker" } : undefined,
    ...decisions.map((decision) => ({
      id: `decision-${decision.id}`,
      title: decision.title,
      meta: decision.decisionMaker ?? decision.owner ?? decision.status ?? "Decision",
    })),
  ]
    .filter((item): item is { id: string; title: string; meta: string } => Boolean(item?.title))
    .filter((item, index, items) => items.findIndex((candidate) => candidate.title === item.title) === index)
    .slice(0, 5);

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

          <div className="exec-position-strip">
            <span className={`exec-status-badge ${programmeTone}`}>{programmeStatusText}</span>
            <p>{programmeReason}</p>
            <div>
              <span>Recovery confidence</span>
              <strong className={`exec-text-${confidenceTone}`}>{executiveToneLabels[confidenceTone]}</strong>
            </div>
          </div>

          <div className="exec-section-label">
            <h3>Executive milestones</h3>
            <p>Highest-level delivery outcomes and confidence in the revised dates.</p>
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
                  <em>{executiveToneLabels[tone]}</em>
                </button>
              );
            })}
          </div>

          <div className="exec-legend" aria-label="Roadmap legend">
            <span><i className="legend-dot green" /> Complete</span>
            <span><i className="legend-dot blue" /> Planned / on track</span>
            <span><i className="legend-dot amber" /> At risk / unconfirmed</span>
            <span><i className="legend-dot red" /> Blocked / overdue</span>
            <span><i className="legend-dot grey" /> Not assessed</span>
            <span><i className="legend-dot executive" /> Executive dependency</span>
          </div>

          <div className="exec-section-label">
            <h3>Executive dependency roadmap</h3>
            <p>Significant schedule dependencies behind each executive milestone.</p>
          </div>
          <div className="exec-pathways">
            {paths.length ? paths.map((path) => (
              <article className={`exec-path ${path.outcome.uid === selectedPath?.outcome.uid ? "selected" : ""}`} key={path.outcome.uid}>
                <h3>{path.outcome.targetMilestone || path.outcome.name}</h3>
                <div className="exec-path-grid">
                  <div className="exec-sequence" aria-label={`${path.outcome.name} dependency pathway`}>
                    {path.dependencies.length ? path.dependencies.map((item) => {
                      const tone = executiveTone(item);
                      return (
                        <div className="exec-node" key={item.uid}>
                          <span>{formatDate(item.finishDate)}</span>
                          <i className={`${tone} ${item.executiveMilestone ? "executive" : ""}`} />
                          <strong>{item.name}</strong>
                        </div>
                      );
                    }) : (
                      <div className="exec-node empty">
                        <span>Not tagged</span>
                        <i className="grey" />
                        <strong>No significant dependencies tagged in the plan yet</strong>
                      </div>
                    )}
                    <ChevronRight className="exec-path-arrow" size={26} />
                  </div>
                  <button className={`exec-outcome-detail ${executiveTone(path.outcome)}`} type="button" onClick={() => setSelectedUid(path.outcome.uid)}>
                    <span>{executiveToneLabels[executiveTone(path.outcome)]}</span>
                    <strong>{path.outcome.name}</strong>
                    <em>{formatDate(path.outcome.finishDate)}</em>
                  </button>
                </div>
              </article>
            )) : (
              <article className="exec-empty-state">
                <h3>No executive milestones found</h3>
                <p>Flag the high-level outcomes in Microsoft Project using the Executive Milestones field, then re-import the XML.</p>
              </article>
            )}
          </div>

          <div className="exec-section-label">
            <h3>Programme watchlist</h3>
            <p>Programme-level items from the schedule and meeting tracker that need attention now.</p>
          </div>
          <div className="exec-watchlist-grid">
            <article className="exec-watch-card">
              <h3>Next key enablers</h3>
              <div className="exec-watch-list gates">
                {keyEnablers.slice(0, 5).map((item) => (
                  <div key={item.uid}>
                    <span className="exec-date-chip">{formatDate(item.finishDate).replace(` ${parseDate(item.finishDate)?.getFullYear() ?? ""}`, "")}</span>
                    <strong>{item.name}</strong>
                  </div>
                ))}
                {!keyEnablers.length ? <p>No executive milestone enablers found in the selected date window.</p> : null}
              </div>
            </article>
            <article className="exec-watch-card">
              <h3>Top risks / issues</h3>
              <div className="exec-watch-list">
                {watchlistItems.map((item) => (
                  <div key={item.id}>
                    <span className="exec-alert-dot amber" />
                    <strong>{item.title}</strong>
                    <em>{item.meta}</em>
                  </div>
                ))}
                {!watchlistItems.length ? <p>No dashboard or red/amber risks or issues are currently flagged.</p> : null}
              </div>
            </article>
            <article className="exec-watch-card">
              <h3>Executive attention</h3>
              <div className="exec-watch-list">
                {attentionItems.map((item) => (
                  <div key={item.id}>
                    <span className="exec-alert-dot red" />
                    <strong>{item.title}</strong>
                    <em>{item.meta}</em>
                  </div>
                ))}
                {!attentionItems.length ? <p>No current asks, blockers or open decisions are currently flagged.</p> : null}
              </div>
            </article>
          </div>
        </section>
      </div>
      {onExportSnapshotPdf ? (
        <div className="snapshot-actions">
          <button className="download-action" type="button" onClick={onExportSnapshotPdf}>
            <Download size={15} />
            Download Executive View PDF
          </button>
        </div>
      ) : null}
    </>
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
  onExportPdf,
  onExportPosterPdf,
  onExportJson,
  onExportSnapshotPdf,
}: {
  schedule: ProgrammeSchedule;
  tracker?: TrackerData;
  dateWindow: DateWindow;
  onExportPdf: () => void;
  onExportPosterPdf: () => void;
  onExportJson: () => void;
  onExportSnapshotPdf: () => void;
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
      meta: "Exports the meeting-ready Executive delivery roadmap.",
      action: "Download Executive View",
      onClick: onExportSnapshotPdf,
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
        <ExecutiveSnapshotView schedule={schedule} tracker={tracker} dateWindow={dateWindow} />
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
}) {
  if (page === "home") return <HomeDashboard schedule={schedule} tracker={tracker} dateWindow={dateWindow} />;
  if (page === "ceo") return <ExecutiveSnapshotView schedule={schedule} tracker={tracker} dateWindow={dateWindow} onExportSnapshotPdf={onExportSnapshotPdf} />;
  if (page === "board") return <BoardReportView schedule={schedule} tracker={tracker} dateWindow={dateWindow} />;
  if (page === "reporting-roadmap") return <ReportingRoadmapView schedule={schedule} dateWindow={dateWindow} selected={selected} onSelect={setSelected} />;
  if (page === "risks") return <RisksIssuesView tracker={tracker} />;
  if (page === "actions") return <ActionsDecisionsView schedule={schedule} tracker={tracker} dateWindow={dateWindow} />;
  if (page === "dependencies") return <DependencyView schedule={schedule} tracker={tracker} />;
  if (page === "workstreams") return <WorkstreamViews schedule={schedule} tracker={tracker} />;
  if (page === "partner") return <PartnerView schedule={schedule} tracker={tracker} />;
  if (page === "downloads") return <DownloadsHub schedule={schedule} tracker={tracker} dateWindow={dateWindow} onExportPdf={onExportPdf} onExportPosterPdf={onExportPosterPdf} onExportJson={onExportJson} onExportSnapshotPdf={onExportSnapshotPdf} />;
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
    } catch (err) {
      setError(err instanceof Error ? err.message : "The file could not be imported.");
    }
  }

  async function importTracker(file: File) {
    setError(undefined);
    try {
      const parsed = await parseMeetingTracker(file);
      setTracker(parsed);
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
      const element = document.getElementById("executive-snapshot-export");
      if (!element) throw new Error("Open the Executive View page or Downloads page before exporting the executive view.");
      await exportElementPdf({
        element,
        title: schedule.title,
        fileNameSuffix: "executive-view",
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "The Executive View PDF could not be generated.");
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
      <SummaryBar schedule={schedule} />
      <nav className="app-nav" aria-label="Application sections">
        {appPages.map((item) => (
          <button type="button" className={`${page === item.key ? "active" : ""} ${item.group}`} onClick={() => setPage(item.key)} key={item.key}>
            {item.key === "workspace" ? <Layers size={15} /> : item.group === "future" ? <CalendarDays size={15} /> : item.key === "workstreams" || item.key === "partner" ? <Users size={15} /> : item.key === "actions" ? <ClipboardCheck size={15} /> : <LayoutDashboard size={15} />}
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
          />
        </section>
      )}
      <DetailDrawer item={selected} onClose={() => setSelected(undefined)} />
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
