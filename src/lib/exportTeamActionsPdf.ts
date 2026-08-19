import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeSchedule } from "../types/programme";
import { formatDate } from "./dateUtils";

type DateWindow = {
  start?: Date;
  end?: Date;
  label: string;
};

export type TeamActionPdfItem = {
  source: string;
  title: string;
  owner?: string;
  stream?: string;
  status?: string;
  displayStatus?: string;
  priority?: string;
  loggedDate?: string;
  dueDate?: string;
  completionDate?: string;
  meetingDate?: string;
  description?: string;
  latestUpdate?: string;
  links?: string;
};

type ExportTeamActionsPdfOptions = {
  schedule: ProgrammeSchedule;
  items: TeamActionPdfItem[];
  dateWindow: DateWindow;
  ownerName?: string;
  ownerPacks?: Array<{ ownerName: string; items: TeamActionPdfItem[] }>;
};

type Rgb = [number, number, number];
type AutoTable = typeof import("jspdf-autotable").default;
type TableRow = string[];

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
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "team-actions";
}

function normaliseText(value?: string): string {
  return (value ?? "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function isComplete(value?: string): boolean {
  return ["complete", "completed", "closed", "done", "resolved", "implemented"].includes(normaliseText(value));
}

function isBlocked(value?: string): boolean {
  const text = normaliseText(value);
  return text.includes("block") || text.includes("hold");
}

function parseDate(value?: string): Date | undefined {
  if (!value) return undefined;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return undefined;
  date.setHours(0, 0, 0, 0);
  return date;
}

function addDays(date: Date, days: number): Date {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function addHeader(doc: JsPDF, schedule: ProgrammeSchedule, dateWindow: DateWindow, subtitle = "Team Actions") {
  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setFillColor(...colours.deep);
  doc.rect(0, 0, pageWidth, 28, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text(`${schedule.title} - ${subtitle}`, 12, 12);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(`Generated ${formatDate(new Date().toISOString())}`, pageWidth - 12, 10, { align: "right" });
  doc.text(`Date window: ${dateWindow.label}`, pageWidth - 12, 17, { align: "right" });
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

function itemTone(item: TeamActionPdfItem): "green" | "amber" | "red" | "blue" {
  const status = item.displayStatus ?? item.status;
  if (isComplete(status)) return "green";
  if (isBlocked(status) || normaliseText(status).includes("overdue") || normaliseText(status).includes("late")) return "red";
  if (normaliseText(status).includes("due soon") || normaliseText(item.priority).includes("high")) return "amber";
  return "blue";
}

function isAttentionItem(item: TeamActionPdfItem): boolean {
  const status = item.displayStatus ?? item.status;
  const due = parseDate(item.dueDate);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return !isComplete(status) && (isBlocked(status) || normaliseText(status).includes("overdue") || normaliseText(status).includes("late") || Boolean(due && due <= addDays(today, 14)));
}

function itemSort(a: TeamActionPdfItem, b: TeamActionPdfItem): number {
  const attentionDelta = Number(isAttentionItem(b)) - Number(isAttentionItem(a));
  if (attentionDelta) return attentionDelta;
  const dueA = parseDate(a.dueDate)?.getTime() ?? Number.POSITIVE_INFINITY;
  const dueB = parseDate(b.dueDate)?.getTime() ?? Number.POSITIVE_INFINITY;
  if (dueA !== dueB) return dueA - dueB;
  return a.title.localeCompare(b.title);
}

function summaryRows(items: TeamActionPdfItem[]): TableRow[] {
  const counts = {
    open: items.filter((item) => !isComplete(item.displayStatus ?? item.status)).length,
    completed: items.filter((item) => isComplete(item.displayStatus ?? item.status)).length,
    late: items.filter((item) => {
      const status = normaliseText(item.displayStatus ?? item.status);
      return status.includes("overdue") || status.includes("late");
    }).length,
    blocked: items.filter((item) => isBlocked(item.displayStatus ?? item.status)).length,
  };
  return [
    ["Open / needs doing", String(counts.open)],
    ["Late", String(counts.late)],
    ["Blocked", String(counts.blocked)],
    ["Completed in view", String(counts.completed)],
    ["Total rows in export", String(items.length)],
  ];
}

function packRows(items: TeamActionPdfItem[]): TableRow[] {
  return items.sort(itemSort).map((item) => [
    item.source,
    item.stream ?? "Not set",
    formatDate(item.loggedDate ?? item.meetingDate),
    formatDate(item.dueDate),
    item.displayStatus ?? item.status ?? "Not set",
    item.title,
    item.latestUpdate || item.description || item.links || "",
  ]);
}

function addPersonPack(
  doc: JsPDF,
  table: AutoTable,
  schedule: ProgrammeSchedule,
  dateWindow: DateWindow,
  ownerName: string,
  items: TeamActionPdfItem[],
  startY = 36,
) {
  addHeader(doc, schedule, dateWindow, `${ownerName} Actions`);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.setTextColor(...colours.ink);
  doc.text(ownerName, 12, startY);
  const activeItems = items.filter((item) => !isComplete(item.displayStatus ?? item.status));
  const dueSoon = activeItems.filter(isAttentionItem).sort(itemSort);
  const upcoming = activeItems.filter((item) => !isAttentionItem(item)).sort(itemSort);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(...colours.muted);
  doc.text(`${dueSoon.length} due soon / attention · ${upcoming.length} upcoming · ${items.length} total assigned rows`, 12, startY + 6);

  table(doc, {
    startY: startY + 12,
    margin: { left: 12, right: 12, bottom: 14 },
    head: [["Source", "Workstream", "Logged", "Due", "Status", "Action / task / milestone", "Notes"]],
    body: dueSoon.length ? packRows(dueSoon) : [["No due soon items", "", "", "", "", "Nothing needs immediate attention.", ""]],
    theme: "grid",
    showHead: "everyPage",
    styles: { font: "helvetica", fontSize: 7.4, cellPadding: 2.2, textColor: colours.ink, lineColor: colours.line, lineWidth: 0.1, overflow: "linebreak", valign: "top" },
    headStyles: { fillColor: colours.amber, textColor: [28, 38, 33], fontStyle: "bold", fontSize: 7.4 },
    alternateRowStyles: { fillColor: colours.pale },
    columnStyles: {
      0: { cellWidth: 21 },
      1: { cellWidth: 26 },
      2: { cellWidth: 18, fontStyle: "bold" },
      3: { cellWidth: 18, fontStyle: "bold" },
      4: { cellWidth: 20, fontStyle: "bold" },
      5: { cellWidth: 51, fontStyle: "bold" },
      6: { cellWidth: 32 },
    },
    didParseCell: (data) => {
      if (data.section !== "body" || data.column.index !== 4) return;
      const item = dueSoon[data.row.index];
      const tone = item ? itemTone(item) : "blue";
      data.cell.styles.textColor = colours[tone];
    },
    didDrawPage: () => addHeader(doc, schedule, dateWindow, `${ownerName} Actions`),
  });

  table(doc, {
    startY: ((doc as JsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? startY + 28) + 8,
    margin: { left: 12, right: 12, bottom: 14 },
    head: [["Source", "Workstream", "Logged", "Due", "Status", "Action / task / milestone", "Notes"]],
    body: upcoming.length ? packRows(upcoming) : [["No upcoming items", "", "", "", "", "No later assigned actions found.", ""]],
    theme: "grid",
    showHead: "everyPage",
    styles: { font: "helvetica", fontSize: 7.4, cellPadding: 2.2, textColor: colours.ink, lineColor: colours.line, lineWidth: 0.1, overflow: "linebreak", valign: "top" },
    headStyles: { fillColor: colours.deep, textColor: [255, 255, 255], fontStyle: "bold", fontSize: 7.4 },
    alternateRowStyles: { fillColor: colours.pale },
    columnStyles: {
      0: { cellWidth: 21 },
      1: { cellWidth: 26 },
      2: { cellWidth: 18, fontStyle: "bold" },
      3: { cellWidth: 18, fontStyle: "bold" },
      4: { cellWidth: 20, fontStyle: "bold" },
      5: { cellWidth: 51, fontStyle: "bold" },
      6: { cellWidth: 32 },
    },
    didDrawPage: () => addHeader(doc, schedule, dateWindow, `${ownerName} Actions`),
  });
}

export async function exportTeamActionsPdf({ schedule, items, dateWindow, ownerName, ownerPacks }: ExportTeamActionsPdfOptions) {
  const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([import("jspdf"), import("jspdf-autotable")]);
  if (ownerPacks?.length || ownerName) {
    const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const table = autoTable as AutoTable;
    const packs = ownerPacks?.length ? ownerPacks : [{ ownerName: ownerName ?? "Team", items }];
    packs.forEach((pack, index) => {
      if (index > 0) doc.addPage();
      addPersonPack(doc, table, schedule, dateWindow, pack.ownerName, pack.items);
    });
    addFooter(doc);
    doc.save(`${fileSlug(schedule.title)}-${ownerPacks?.length ? "team-action-packs" : fileSlug(ownerName ?? "team") + "-actions"}.pdf`);
    return;
  }

  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  addHeader(doc, schedule, dateWindow);

  const table = autoTable as AutoTable;
  table(doc, {
    startY: 36,
    margin: { left: 12, right: 12 },
    head: [["Measure", "Count"]],
    body: summaryRows(items),
    theme: "grid",
    styles: { font: "helvetica", fontSize: 8, cellPadding: 2.4, textColor: colours.ink, lineColor: colours.line, lineWidth: 0.1 },
    headStyles: { fillColor: colours.deep, textColor: [255, 255, 255], fontStyle: "bold" },
    columnStyles: {
      0: { cellWidth: 44, fontStyle: "bold" },
      1: { cellWidth: 20, halign: "right" },
    },
  });

  const rows = items.map((item) => [
    item.source,
    item.owner ?? "No owner",
    item.stream ?? "Not set",
    formatDate(item.loggedDate ?? item.meetingDate),
    formatDate(item.dueDate),
    item.displayStatus ?? item.status ?? "Not set",
    item.title,
    item.latestUpdate || item.description || item.links || "",
  ]);

  table(doc, {
    startY: ((doc as JsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 62) + 8,
    margin: { left: 12, right: 12, bottom: 14 },
    head: [["Source", "Owner", "Workstream", "Logged", "Due", "Status", "Action / task / milestone", "Notes"]],
    body: rows.length ? rows : [["No rows", "", "", "", "", "", "No actions matched the selected filters.", ""]],
    theme: "grid",
    showHead: "everyPage",
    styles: { font: "helvetica", fontSize: 7.1, cellPadding: 2.3, textColor: colours.ink, lineColor: colours.line, lineWidth: 0.1, overflow: "linebreak", valign: "top" },
    headStyles: { fillColor: colours.deep, textColor: [255, 255, 255], fontStyle: "bold", fontSize: 7.2 },
    alternateRowStyles: { fillColor: colours.pale },
    columnStyles: {
      0: { cellWidth: 24 },
      1: { cellWidth: 28 },
      2: { cellWidth: 34 },
      3: { cellWidth: 20, fontStyle: "bold" },
      4: { cellWidth: 20, fontStyle: "bold" },
      5: { cellWidth: 24, fontStyle: "bold" },
      6: { cellWidth: 58, fontStyle: "bold" },
      7: { cellWidth: 69 },
    },
    didParseCell: (data) => {
      if (data.section !== "body" || data.column.index !== 5) return;
      const item = items[data.row.index];
      const tone = item ? itemTone(item) : "blue";
      data.cell.styles.textColor = colours[tone];
    },
    didDrawPage: () => addHeader(doc, schedule, dateWindow),
  });

  addFooter(doc);
  doc.save(`${fileSlug(schedule.title)}-team-actions-a4.pdf`);
}
