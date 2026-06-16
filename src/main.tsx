import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  CalendarDays,
  ChevronRight,
  Diamond,
  Download,
  FileUp,
  Filter,
  GitBranch,
  Layers,
  Milestone,
  Search,
  ShieldCheck,
  SlidersHorizontal,
  X,
} from "lucide-react";
import { exportProgrammePdf } from "./lib/exportProgrammePdf";
import { parseMicrosoftProjectXml } from "./lib/parseMicrosoftProjectXml";
import { clamp, formatDate, parseDate, uniqueSorted } from "./lib/dateUtils";
import type { ProgrammeFilters, ProgrammeItem, ProgrammeSchedule, ProgrammeView } from "./types/programme";
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
  const [filters, setFilters] = useState<ProgrammeFilters>(initialFilters);
  const [view, setView] = useState<ProgrammeView>("roadmap");
  const [selected, setSelected] = useState<ProgrammeItem | undefined>();
  const [error, setError] = useState<string | undefined>();
  const [baselineNumber, setBaselineNumber] = useState(3);
  const [sourceXml, setSourceXml] = useState<{ xml: string; fileName: string } | undefined>();

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
        baselineNumber,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "The PDF report could not be generated.");
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
        <label className="upload-button">
          <FileUp size={18} />
          Import Project XML
          <input type="file" accept=".xml,text/xml,application/xml" onChange={(event) => event.target.files?.[0] && importFile(event.target.files[0])} />
        </label>
      </header>

      {error ? <div className="error">{error}</div> : null}
      <SummaryBar schedule={schedule} />

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
              <button type="button" onClick={exportJson}><Download size={15} /> JSON</button>
            </div>
          </div>
          <Timeline schedule={schedule} items={visibleItems} selected={selected} onSelect={setSelected} dateWindow={dateWindow} />
          <InsightsPanel schedule={schedule} />
        </section>
      </div>
      <DetailDrawer item={selected} onClose={() => setSelected(undefined)} />
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
