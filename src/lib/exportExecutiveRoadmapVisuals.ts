import type { jsPDF as JsPDF } from "jspdf";
import type { ProgrammeItem, ProgrammeSchedule } from "../types/programme";
import type { TrackerData } from "../types/reporting";
import { formatDate } from "./dateUtils";
import {
  buildExecutiveRoadmapModel,
  executiveTone,
  executiveToneAssessment,
  executiveToneLabel,
  executiveToneLabels,
  type DateWindow,
  type ExecutiveTone,
} from "./executiveRoadmapData";

type Rgb = [number, number, number];

const colours: Record<"ink" | "muted" | "deep" | "line" | "pale" | "green" | "amber" | "red" | "blue" | "grey" | "white", Rgb> = {
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
  white: [255, 255, 255],
};

const toneColours: Record<ExecutiveTone, Rgb> = {
  green: colours.green,
  blue: colours.blue,
  amber: colours.amber,
  red: colours.red,
  grey: colours.grey,
};

const toneFills: Record<ExecutiveTone, Rgb> = {
  green: [231, 245, 237],
  blue: [232, 242, 251],
  amber: [255, 240, 210],
  red: [255, 231, 229],
  grey: [237, 241, 239],
};

function statusTone(value?: string): ExecutiveTone {
  const text = (value ?? "").toLowerCase();
  if (text.includes("red")) return "red";
  if (text.includes("amber")) return "amber";
  if (text.includes("green")) return "green";
  return "grey";
}

function fileSlug(value: string): string {
  return value.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "executive-roadmap";
}

function escapeHtml(value?: string): string {
  return (value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function saveTextFile(content: string, filename: string, type: string) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function setText(doc: JsPDF, colour: Rgb) {
  doc.setTextColor(colour[0], colour[1], colour[2]);
}

function setFill(doc: JsPDF, colour: Rgb) {
  doc.setFillColor(colour[0], colour[1], colour[2]);
}

function setDraw(doc: JsPDF, colour: Rgb) {
  doc.setDrawColor(colour[0], colour[1], colour[2]);
}

function addPosterHeader(doc: JsPDF, schedule: ProgrammeSchedule, reportDate: string, subtitle: string) {
  const pageWidth = doc.internal.pageSize.getWidth();
  setFill(doc, colours.deep);
  doc.rect(0, 0, pageWidth, 32, "F");
  setText(doc, colours.white);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("DAF Executive Delivery Roadmap", 14, 12);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text(schedule.title, 14, 22);
  doc.text(subtitle, pageWidth - 14, 12, { align: "right" });
  doc.text(`Generated ${formatDate(reportDate)}`, pageWidth - 14, 22, { align: "right" });
  setText(doc, colours.ink);
}

function addPosterFooter(doc: JsPDF) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const pageCount = doc.getNumberOfPages();
  for (let page = 1; page <= pageCount; page += 1) {
    doc.setPage(page);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    setText(doc, colours.muted);
    doc.text(`Page ${page} of ${pageCount}`, pageWidth - 14, pageHeight - 8, { align: "right" });
  }
}

function drawSummaryCard(doc: JsPDF, x: number, y: number, w: number, title: string, value: string, tone?: ExecutiveTone) {
  setDraw(doc, tone ? toneColours[tone] : colours.line);
  setFill(doc, tone ? toneFills[tone] : colours.white);
  doc.roundedRect(x, y, w, 25, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.4);
  setText(doc, colours.muted);
  doc.text(title.toUpperCase(), x + 4, y + 7);
  doc.setFontSize(9.2);
  setText(doc, colours.ink);
  doc.text(doc.splitTextToSize(value, w - 8).slice(0, 2), x + 4, y + 15);
}

function drawLegend(doc: JsPDF, x: number, y: number, w: number): number {
  const items: Array<{ tone: ExecutiveTone; label: string }> = [
    { tone: "green", label: "Complete / confirmed" },
    { tone: "blue", label: "Planned / dated" },
    { tone: "amber", label: "Date assumption / not confirmed" },
    { tone: "red", label: "Blocked / overdue" },
    { tone: "grey", label: "Not assessed" },
  ];
  setDraw(doc, colours.line);
  setFill(doc, colours.pale);
  doc.roundedRect(x, y, w, 18, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.2);
  setText(doc, colours.muted);
  doc.text("Colour status", x + 4, y + 7);
  let cursor = x + 34;
  items.forEach((item) => {
    setFill(doc, toneColours[item.tone]);
    doc.circle(cursor, y + 6, 2.4, "F");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.1);
    setText(doc, colours.ink);
    doc.text(item.label, cursor + 5, y + 8);
    cursor += doc.getTextWidth(item.label) + 18;
  });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7);
  setText(doc, colours.muted);
  doc.text("Orange means Date Assumption is Yes or the source RAG is Amber. Decision and dependency flags are shown as detail evidence.", x + 4, y + 15);
  return y + 24;
}

function drawMilestoneCard(doc: JsPDF, item: ProgrammeItem, x: number, y: number, w: number, h: number) {
  const tone = executiveTone(item);
  const assessment = executiveToneAssessment(item);
  setDraw(doc, toneColours[tone]);
  setFill(doc, toneFills[tone]);
  doc.roundedRect(x, y, w, h, 2.5, 2.5, "FD");
  setFill(doc, toneColours[tone]);
  doc.rect(x, y, w, 4, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8.5);
  setText(doc, colours.muted);
  doc.text(formatDate(item.finishDate), x + 4, y + 12);
  doc.setFontSize(9.4);
  setText(doc, colours.ink);
  doc.text(doc.splitTextToSize(item.name, w - 8).slice(0, 3), x + 4, y + 22);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(6.6);
  setText(doc, colours.muted);
  doc.text(doc.splitTextToSize(assessment.summary, w - 8).slice(0, 3), x + 4, y + h - 21);
  setFill(doc, toneColours[tone]);
  doc.roundedRect(x + 4, y + h - 12, 22, 7, 3, 3, "F");
  doc.setFontSize(6.2);
  setText(doc, colours.white);
  doc.text(executiveToneLabel(item), x + 15, y + h - 7.2, { align: "center" });
}

function drawPathRow(doc: JsPDF, title: string, items: ProgrammeItem[], y: number): number {
  const margin = 14;
  const pageWidth = doc.internal.pageSize.getWidth();
  const width = pageWidth - margin * 2;
  const rowHeight = 34;
  setDraw(doc, colours.line);
  setFill(doc, colours.pale);
  doc.roundedRect(margin, y, width, rowHeight, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8.8);
  setText(doc, colours.ink);
  doc.text(doc.splitTextToSize(title, 58).slice(0, 2), margin + 4, y + 10);
  const pathX = margin + 66;
  const pathWidth = width - 72;
  const chipWidth = Math.min(42, pathWidth / Math.max(1, items.length) - 4);
  const gap = items.length > 1 ? (pathWidth - chipWidth * items.length) / (items.length - 1) : 0;
  setDraw(doc, colours.line);
  doc.line(pathX, y + 16, pathX + pathWidth, y + 16);
  items.forEach((item, index) => {
    const x = pathX + index * (chipWidth + gap);
    const tone = executiveTone(item);
    setFill(doc, toneColours[tone]);
    doc.circle(x + 3, y + 16, 2.8, "F");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(6.4);
    setText(doc, colours.muted);
    doc.text(formatDate(item.finishDate).replace(" 20", " "), x, y + 8);
    doc.setFontSize(6.8);
    setText(doc, colours.ink);
    doc.text(doc.splitTextToSize(item.name, chipWidth).slice(0, 3), x, y + 23);
  });
  return y + rowHeight + 7;
}

export async function exportExecutiveRoadmapImage(element: HTMLElement, schedule: ProgrammeSchedule) {
  const { default: html2canvas } = await import("html2canvas");
  document.body.classList.add("executive-export-light");
  document.body.classList.add("snapshot-exporting");
  try {
    const width = Math.max(1180, element.scrollWidth, element.offsetWidth);
    const height = Math.max(element.scrollHeight, element.offsetHeight);
    const maxCanvasArea = 32_000_000;
    const preferredScale = 2;
    const scale = Math.min(preferredScale, Math.sqrt(maxCanvasArea / Math.max(1, width * height)));
    const canvas = await html2canvas(element, {
      backgroundColor: "#ffffff",
      scale: Math.max(1, scale),
      useCORS: true,
      windowWidth: Math.max(1440, width),
      width,
      height,
      scrollX: 0,
      scrollY: 0,
      onclone: (clonedDocument) => {
        clonedDocument.body.classList.add("executive-export-light", "snapshot-exporting");
      },
    });
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
    if (!blob) throw new Error("The roadmap image was too large to render. Try a shorter date window or export the poster PDF.");
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${fileSlug(schedule.title)}-executive-roadmap-snapshot.png`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } finally {
    document.body.classList.remove("executive-export-light");
    document.body.classList.remove("snapshot-exporting");
  }
}

export async function exportExecutiveRoadmapPosterPdf(schedule: ProgrammeSchedule, tracker: TrackerData | undefined, dateWindow: DateWindow) {
  const { default: jsPDF } = await import("jspdf");
  const model = buildExecutiveRoadmapModel(schedule, tracker, dateWindow);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
  addPosterHeader(doc, schedule, model.reportDate, "Executive roadmap poster");

  let y = 44;
  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 14;
  const cardGap = 6;
  const summaryWidth = (pageWidth - margin * 2 - cardGap * 3) / 4;
  drawSummaryCard(doc, margin, y, summaryWidth, "Programme status", model.programmeStatus, statusTone(model.programmeStatus));
  drawSummaryCard(doc, margin + (summaryWidth + cardGap), y, summaryWidth, "Original delivery plan", formatDate(model.originalDeliveryDate));
  drawSummaryCard(doc, margin + (summaryWidth + cardGap) * 2, y, summaryWidth, "Current forecast", formatDate(model.currentDeliveryDate));
  drawSummaryCard(doc, margin + (summaryWidth + cardGap) * 3, y, summaryWidth, "Forecast basis", model.forecastBasis);
  y += 32;
  y = drawLegend(doc, margin, y, pageWidth - margin * 2);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  setText(doc, colours.ink);
  doc.text("Executive milestones", margin, y);
  y += 8;
  const milestoneWidth = (pageWidth - margin * 2 - cardGap * Math.max(0, model.outcomes.length - 1)) / Math.max(1, model.outcomes.length);
  model.outcomes.forEach((item, index) => {
    drawMilestoneCard(doc, item, margin + index * (milestoneWidth + cardGap), y, milestoneWidth, 62);
  });
  y += 78;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  setText(doc, colours.ink);
  doc.text("Dependency roadmap", margin, y);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  setText(doc, colours.muted);
  doc.text("Key stream milestones and linked cross-stream milestone dependencies from the imported Project plan.", margin, y + 6);
  y += 14;
  model.paths.forEach((path) => {
    const needed = 41;
    if (y + needed > doc.internal.pageSize.getHeight() - 18) {
      doc.addPage();
      addPosterHeader(doc, schedule, model.reportDate, "Executive roadmap poster");
      y = 44;
    }
    y = drawPathRow(doc, path.outcome.name, [...path.dependencies, path.outcome], y);
  });

  addPosterFooter(doc);
  doc.save(`${fileSlug(schedule.title)}-executive-roadmap-poster.pdf`);
}

export function exportExecutiveRoadmapHtml(schedule: ProgrammeSchedule, tracker: TrackerData | undefined, dateWindow: DateWindow) {
  const model = buildExecutiveRoadmapModel(schedule, tracker, dateWindow);
  const milestoneCards = model.outcomes.map((item) => {
    const tone = executiveTone(item);
    const assessment = executiveToneAssessment(item);
    return `<article class="milestone ${tone}">
      <span>${escapeHtml(formatDate(item.finishDate))}</span>
      <strong>${escapeHtml(item.name)}</strong>
      <p>${escapeHtml(assessment.summary)}</p>
      <em>${escapeHtml(executiveToneLabel(item))}</em>
    </article>`;
  }).join("");
  const pathRows = model.paths.map((path) => `
    <section class="path">
      <h2>${escapeHtml(path.outcome.name)}</h2>
      <div class="sequence">
        ${[...path.dependencies, path.outcome].map((item) => {
          const tone = executiveTone(item);
          const assessment = executiveToneAssessment(item);
          return `<article class="node ${tone}">
            <span>${escapeHtml(formatDate(item.finishDate))}</span>
            <i></i>
            <strong>${escapeHtml(item.name)}</strong>
            <p>${escapeHtml(assessment.summary)}</p>
          </article>`;
        }).join("")}
      </div>
    </section>
  `).join("");
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(schedule.title)} - Executive Roadmap</title>
  <style>
    :root { font-family: Inter, Arial, sans-serif; color: #17201b; background: #f4f7f5; }
    body { margin: 0; }
    main { max-width: 1440px; margin: 0 auto; padding: 28px; }
    header { display: flex; justify-content: space-between; gap: 24px; align-items: center; margin: -28px -28px 26px; padding: 28px; color: white; background: #214c43; }
    h1 { margin: 0 0 6px; font-size: 32px; }
    p { margin: 0; color: #647269; }
    header p { color: #eaf4ee; }
    .summary, .milestones { display: grid; gap: 12px; }
    .summary { grid-template-columns: repeat(4, 1fr); margin-bottom: 24px; }
    .summary article, .milestone, .path, .legend { border: 1px solid #c7d1cb; border-radius: 8px; background: white; box-shadow: 0 10px 24px rgba(23,32,27,.08); }
    .summary article { padding: 16px; }
    .summary span { display: block; color: #647269; font-size: 12px; font-weight: 800; text-transform: uppercase; }
    .summary strong { display: block; margin-top: 8px; font-size: 18px; }
    .legend { display: flex; flex-wrap: wrap; gap: 12px 22px; align-items: center; margin-bottom: 24px; padding: 14px 16px; }
    .legend strong { margin-right: 8px; }
    .legend span { display: inline-flex; align-items: center; gap: 8px; color: #647269; font-size: 14px; font-weight: 800; }
    .legend i { width: 14px; height: 14px; border-radius: 999px; background: var(--tone); }
    .legend p { flex-basis: 100%; font-size: 13px; }
    .milestones { grid-template-columns: repeat(${Math.max(1, model.outcomes.length)}, 1fr); margin-bottom: 28px; }
    .summary article.status, .milestone { border-color: var(--tone); background: var(--card-bg); }
    .milestone { min-height: 185px; padding: 18px; border-top: 7px solid var(--tone); }
    .milestone span, .node span { color: #647269; font-weight: 800; }
    .milestone strong { display: block; margin: 20px 0 10px; font-size: 20px; line-height: 1.35; }
    .milestone p, .node p { margin: 0 0 12px; color: #647269; font-size: 13px; line-height: 1.35; }
    .milestone em { display: inline-block; padding: 7px 10px; border-radius: 999px; color: white; background: var(--tone); font-style: normal; font-weight: 900; }
    .green { --tone: #2e7d55; --card-bg: #e7f5ed; } .blue { --tone: #3d78a9; --card-bg: #e8f2fb; } .amber { --tone: #ff8a00; --card-bg: #fff0d2; } .red { --tone: #b33a32; --card-bg: #ffe7e5; } .grey { --tone: #7e8c84; --card-bg: #edf1ef; }
    .path { margin-bottom: 16px; padding: 18px; break-inside: avoid; }
    .path h2 { margin: 0 0 16px; font-size: 20px; }
    .sequence { display: grid; grid-auto-flow: column; grid-auto-columns: minmax(130px, 1fr); gap: 14px; align-items: start; border-top: 4px solid #c7d1cb; padding-top: 14px; overflow-x: auto; }
    .node { display: grid; gap: 8px; text-align: center; }
    .node i { justify-self: center; width: 18px; height: 18px; border-radius: 999px; background: var(--tone); }
    .node strong { font-size: 14px; line-height: 1.3; }
    @media print { body { background: white; } main { padding: 12mm; } header { margin: -12mm -12mm 10mm; } .path, .milestone, .summary article, .legend { box-shadow: none; } }
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>DAF Executive Delivery Roadmap</h1>
        <p>${escapeHtml(schedule.title)}</p>
      </div>
      <strong>Generated ${escapeHtml(formatDate(model.reportDate))}</strong>
    </header>
    <section class="summary">
      <article class="status ${statusTone(model.programmeStatus)}"><span>Programme status</span><strong>${escapeHtml(model.programmeStatus)}</strong></article>
      <article><span>Original delivery plan</span><strong>${escapeHtml(formatDate(model.originalDeliveryDate))}</strong></article>
      <article><span>Current forecast</span><strong>${escapeHtml(formatDate(model.currentDeliveryDate))}</strong></article>
      <article><span>Forecast basis</span><strong>${escapeHtml(model.forecastBasis)}</strong></article>
    </section>
    <section class="legend" aria-label="Colour status legend">
      <strong>Colour status</strong>
      <span class="green"><i></i>Complete / confirmed</span>
      <span class="blue"><i></i>Planned / dated</span>
      <span class="amber"><i></i>Date assumption / not confirmed</span>
      <span class="red"><i></i>Blocked / overdue</span>
      <span class="grey"><i></i>Not assessed</span>
      <p>Orange means Date Assumption is Yes or the source RAG is Amber. Decision and dependency flags are shown as detail evidence.</p>
    </section>
    <section class="milestones">${milestoneCards}</section>
    ${pathRows}
  </main>
</body>
</html>`;
  saveTextFile(html, `${fileSlug(schedule.title)}-executive-roadmap.html`, "text/html;charset=utf-8");
}
