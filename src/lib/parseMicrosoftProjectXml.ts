import type {
  ProgrammeBaseline,
  ProgrammeDependency,
  ProgrammeItem,
  ProgrammeResource,
  ProgrammeSchedule,
  ProgrammeStatus,
} from "../types/programme";
import { daysBetween, parseDate } from "./dateUtils";

const CUSTOM_FIELDS: Record<string, keyof ProgrammeItem> = {
  "188743731": "stream",
  "188743734": "roadmapMilestone",
  "188743737": "milestoneType",
  "188743740": "approvalBody",
  "188743743": "version",
  "188743746": "visibility",
  "188743747": "roadmapView",
  "188743752": "discussed",
};

function childText(node: Element, tag: string): string | undefined {
  return node.getElementsByTagName(tag)[0]?.textContent?.trim() || undefined;
}

function directChildren(node: Element | Document, tag: string): Element[] {
  return Array.from(node.getElementsByTagName(tag)).filter((child) => child.parentElement === node);
}

function asBool(value?: string): boolean {
  return value === "1" || value?.toLowerCase() === "true" || value?.toLowerCase() === "yes";
}

function asNumber(value?: string): number | undefined {
  if (!value) return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function resolveLookupValue(value?: string, valueGuid?: string, lookupByGuid?: Map<string, string>): string | undefined {
  if (!value) return undefined;
  if (valueGuid && lookupByGuid?.has(valueGuid)) return lookupByGuid.get(valueGuid);
  return value;
}

function statusFor(item: Pick<ProgrammeItem, "percentComplete" | "finishDate" | "isCritical" | "delayDays">, statusDate?: string): ProgrammeStatus {
  if ((item.percentComplete ?? 0) >= 100) return "complete";
  const status = parseDate(statusDate);
  const finish = parseDate(item.finishDate);
  if (item.delayDays && item.delayDays > 10) return "late";
  if (item.delayDays && item.delayDays > 0) return "at-risk";
  if (status && finish && finish < status) return "late";
  if ((item.percentComplete ?? 0) > 0) return "in-progress";
  if (item.isCritical) return "at-risk";
  return "not-started";
}

function buildLookup(root: Document): Map<string, string> {
  const lookup = new Map<string, string>();
  Array.from(root.getElementsByTagName("OutlineCode")).forEach((outline) => {
    Array.from(outline.getElementsByTagName("Value")).forEach((valueNode) => {
      const guid = childText(valueNode, "FieldGUID");
      const label = childText(valueNode, "Value");
      const id = childText(valueNode, "ValueID");
      if (guid && label) lookup.set(guid, label);
      if (id && label) lookup.set(id, label);
    });
  });
  return lookup;
}

function parseBaselines(task: Element): ProgrammeBaseline[] {
  return directChildren(task, "Baseline").map((baseline) => ({
    number: asNumber(childText(baseline, "Number")) ?? 0,
    start: childText(baseline, "Start"),
    finish: childText(baseline, "Finish"),
    duration: childText(baseline, "Duration"),
    work: childText(baseline, "Work"),
  }));
}

function parsePredecessors(task: Element): ProgrammeDependency[] {
  return directChildren(task, "PredecessorLink").map((link) => ({
    predecessorUid: childText(link, "PredecessorUID"),
    type: childText(link, "Type"),
    lag: asNumber(childText(link, "LinkLag")),
  }));
}

function parseResources(root: Document): ProgrammeResource[] {
  const resourcesNode = root.getElementsByTagName("Resources")[0];
  if (!resourcesNode) return [];
  return directChildren(resourcesNode, "Resource")
    .map((resource) => ({
      uid: childText(resource, "UID") ?? "",
      name: childText(resource, "Name") ?? "Unnamed resource",
      initials: childText(resource, "Initials"),
      email: childText(resource, "EmailAddress"),
    }))
    .filter((resource) => resource.uid && resource.name !== "Unnamed resource");
}

function buildAssignments(root: Document, resources: ProgrammeResource[]): Map<string, string[]> {
  const resourceByUid = new Map(resources.map((resource) => [resource.uid, resource.name]));
  const assignments = new Map<string, string[]>();
  const assignmentsNode = root.getElementsByTagName("Assignments")[0];
  if (!assignmentsNode) return assignments;
  directChildren(assignmentsNode, "Assignment").forEach((assignment) => {
    const taskUid = childText(assignment, "TaskUID");
    const resourceUid = childText(assignment, "ResourceUID");
    const resourceName = resourceUid ? resourceByUid.get(resourceUid) : undefined;
    if (!taskUid || !resourceName) return;
    assignments.set(taskUid, [...(assignments.get(taskUid) ?? []), resourceName]);
  });
  return assignments;
}

function applyCustomFields(task: Element, item: ProgrammeItem, lookup: Map<string, string>): ProgrammeItem {
  directChildren(task, "ExtendedAttribute").forEach((attribute) => {
    const fieldId = childText(attribute, "FieldID");
    const fieldName = CUSTOM_FIELDS[fieldId ?? ""];
    if (!fieldName) return;
    const raw = resolveLookupValue(childText(attribute, "Value"), childText(attribute, "ValueGUID"), lookup);
    if (fieldName === "roadmapMilestone") item.roadmapMilestone = raw?.toLowerCase() === "yes" || asBool(raw);
    else if (fieldName === "discussed") item.discussed = asBool(raw);
    else if (raw && raw !== "None" && raw !== "N/A") (item[fieldName] as string | boolean | undefined) = raw;
  });
  return item;
}

function buildHierarchy(items: ProgrammeItem[]): ProgrammeItem[] {
  const stack: ProgrammeItem[] = [];
  items.forEach((item) => {
    while (stack.length && stack[stack.length - 1].outlineLevel >= item.outlineLevel) stack.pop();
    const parent = stack[stack.length - 1];
    if (parent) {
      item.parentUid = parent.uid;
      parent.childUids = [...(parent.childUids ?? []), item.uid];
    }
    stack.push(item);
  });
  return items;
}

function buildSuccessors(items: ProgrammeItem[]): ProgrammeItem[] {
  const byUid = new Map(items.map((item) => [item.uid, item]));
  items.forEach((item) => {
    item.predecessors.forEach((dependency) => {
      if (!dependency.predecessorUid) return;
      const predecessor = byUid.get(dependency.predecessorUid);
      predecessor?.successors.push({
        successorUid: item.uid,
        type: dependency.type,
        lag: dependency.lag,
      });
    });
  });
  return items;
}

export function parseMicrosoftProjectXml(xml: string, sourceFileName?: string, baselineNumber = 3): ProgrammeSchedule {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const parserError = doc.getElementsByTagName("parsererror")[0];
  if (parserError) throw new Error(parserError.textContent ?? "The XML could not be parsed.");

  const lookup = buildLookup(doc);
  const resources = parseResources(doc);
  const assignments = buildAssignments(doc, resources);
  const tasksNode = doc.getElementsByTagName("Tasks")[0];
  if (!tasksNode) throw new Error("No Tasks node was found in this Microsoft Project XML file.");

  const statusDate = childText(doc.documentElement, "StatusDate") ?? childText(doc.documentElement, "CurrentDate");
  const items = directChildren(tasksNode, "Task")
    .filter((task) => childText(task, "UID") && childText(task, "Name") && !asBool(childText(task, "IsNull")))
    .map((task) => {
      const uid = childText(task, "UID") ?? "";
      const isSummary = asBool(childText(task, "Summary"));
      const isMilestone = asBool(childText(task, "Milestone"));
      const baselines = parseBaselines(task);
      const selectedBaseline = baselines.find((baseline) => baseline.number === baselineNumber) ?? baselines[0];
      const finishDate = childText(task, "Finish");
      const delayDays = selectedBaseline?.finish ? Math.max(0, daysBetween(selectedBaseline.finish, finishDate) ?? 0) : undefined;
      const item: ProgrammeItem = {
        uid,
        id: childText(task, "ID"),
        name: childText(task, "Name") ?? "Unnamed task",
        wbs: childText(task, "WBS"),
        outlineNumber: childText(task, "OutlineNumber"),
        outlineLevel: asNumber(childText(task, "OutlineLevel")) ?? 1,
        itemType: isSummary ? "summary" : isMilestone ? "milestone" : "task",
        isSummary,
        isMilestone,
        isCritical: asBool(childText(task, "Critical")),
        isActive: !childText(task, "Active") || asBool(childText(task, "Active")),
        startDate: childText(task, "Start"),
        finishDate,
        duration: childText(task, "Duration"),
        remainingDuration: childText(task, "RemainingDuration"),
        percentComplete: asNumber(childText(task, "PercentComplete")),
        percentWorkComplete: asNumber(childText(task, "PercentWorkComplete")),
        earlyStart: childText(task, "EarlyStart"),
        earlyFinish: childText(task, "EarlyFinish"),
        lateStart: childText(task, "LateStart"),
        lateFinish: childText(task, "LateFinish"),
        freeSlack: asNumber(childText(task, "FreeSlack")),
        totalSlack: asNumber(childText(task, "TotalSlack")),
        startVariance: asNumber(childText(task, "StartVariance")),
        finishVariance: asNumber(childText(task, "FinishVariance")),
        baselineStart: selectedBaseline?.start,
        baselineFinish: selectedBaseline?.finish,
        baselineNumber: selectedBaseline?.number,
        baselines,
        predecessors: parsePredecessors(task),
        successors: [],
        resourceNames: assignments.get(uid),
        roadmapMilestone: false,
        delayDays,
        status: "future",
      };
      applyCustomFields(task, item, lookup);
      item.status = statusFor(item, statusDate);
      return item;
    });

  const normalisedItems = buildSuccessors(buildHierarchy(items));
  return {
    title: childText(doc.documentElement, "Title") ?? childText(doc.documentElement, "Name") ?? "Imported programme",
    name: childText(doc.documentElement, "Name"),
    startDate: childText(doc.documentElement, "StartDate"),
    finishDate: childText(doc.documentElement, "FinishDate"),
    statusDate,
    currentDate: childText(doc.documentElement, "CurrentDate"),
    multipleCriticalPaths: asBool(childText(doc.documentElement, "MultipleCriticalPaths")),
    items: normalisedItems,
    resources,
    importedAt: new Date().toISOString(),
    sourceFileName,
  };
}
