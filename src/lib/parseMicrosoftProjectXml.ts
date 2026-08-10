import type {
  ProgrammeBaseline,
  ProgrammeDependency,
  ProgrammeItem,
  ProgrammeResource,
  ProgrammeSchedule,
  ProgrammeStatus,
} from "../types/programme";
import { daysBetween, parseDate } from "./dateUtils";

const FALLBACK_CUSTOM_FIELDS: Record<string, keyof ProgrammeItem> = {
  "188743731": "stream",
  "188743734": "roadmapMilestone",
  "188743737": "milestoneType",
  "188743740": "approvalBody",
  "188743743": "version",
  "188743746": "visibility",
  "188743747": "roadmapView",
  "188743748": "milestoneLevel",
  "188743749": "dependencyLevel",
  "188743750": "criticalPathReview",
  "188743752": "discussed",
  "188743753": "executiveMilestone",
  "188743754": "boardReportable",
  "188743755": "decisionRequired",
  "188743756": "externalDependency",
  "188743758": "dependencyAnchor",
  "188743759": "governanceGate",
  "188743999": "ragStatus",
  "188744006": "targetMilestone",
  "188744007": "workstreamAccountableOwner",
  "188744008": "deliverySupportRoles",
  "188744009": "projectManagerAssurance",
};

function childText(node: Element, tag: string): string | undefined {
  return node.getElementsByTagName(tag)[0]?.textContent?.trim() || undefined;
}

function directChildren(node: Element | Document, tag: string): Element[] {
  return Array.from(node.getElementsByTagName(tag)).filter((child) => child.parentElement === node);
}

function projectText(root: Element, tag: string): string | undefined {
  return directChildren(root, tag)[0]?.textContent?.trim() || undefined;
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
  if (lookupByGuid?.has(value)) return lookupByGuid.get(value);
  return value;
}

function normaliseFieldLabel(value?: string): string {
  return (value ?? "")
    .toLowerCase()
    .replace(/\([^)]*\)/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function customFieldForDefinition(fieldName?: string, alias?: string): keyof ProgrammeItem | undefined {
  const labels = [normaliseFieldLabel(alias), normaliseFieldLabel(fieldName)];
  if (labels.some((label) => label === "stream" || label === "text1")) return "stream";
  if (labels.some((label) => label === "roadmap milestone" || label === "text2")) return "roadmapMilestone";
  if (labels.some((label) => label === "milestone type" || label === "text3")) return "milestoneType";
  if (labels.some((label) => label === "approval body" || label === "text4")) return "approvalBody";
  if (labels.some((label) => label === "version" || label === "text5")) return "version";
  if (labels.some((label) => label === "visibility" || label === "text6")) return "visibility";
  if (labels.some((label) => label === "roadmap view" || label === "text7")) return "roadmapView";
  if (labels.some((label) => label === "milestone level" || label === "text8")) return "milestoneLevel";
  if (labels.some((label) => label === "dependency level" || label === "text9")) return "dependencyLevel";
  if (labels.some((label) => label === "critical path review" || label === "text10")) return "criticalPathReview";
  if (labels.some((label) => label === "rag status" || label === "text13")) return "ragStatus";
  if (labels.some((label) => label === "date confidence" || label === "text19")) return "dateConfidence";
  if (labels.some((label) => label === "target milestone" || label === "text20")) return "targetMilestone";
  if (labels.some((label) => label === "workstream accountable owner" || label === "text21")) return "workstreamAccountableOwner";
  if (labels.some((label) => label === "delivery support roles" || label === "text22")) return "deliverySupportRoles";
  if (labels.some((label) => label === "project manager assurance" || label === "text23")) return "projectManagerAssurance";
  if (labels.some((label) => label === "discussed" || label === "flag1")) return "discussed";
  if (labels.some((label) => label === "executive milestones" || label === "flag2")) return "executiveMilestone";
  if (labels.some((label) => label === "board reportable" || label === "flag3")) return "boardReportable";
  if (labels.some((label) => label === "decision required" || label === "flag4")) return "decisionRequired";
  if (labels.some((label) => label === "external dependency" || label === "flag5")) return "externalDependency";
  if (labels.some((label) => label === "dependency anchor" || label === "flag7")) return "dependencyAnchor";
  if (labels.some((label) => label === "governance gate" || label === "flag8")) return "governanceGate";
  return undefined;
}

function buildCustomFieldMap(root: Document): Map<string, keyof ProgrammeItem> {
  const fields = new Map<string, keyof ProgrammeItem>(Object.entries(FALLBACK_CUSTOM_FIELDS));
  const definitionsNode = root.getElementsByTagName("ExtendedAttributes")[0];
  if (!definitionsNode) return fields;

  directChildren(definitionsNode, "ExtendedAttribute").forEach((definition) => {
    const fieldId = childText(definition, "FieldID");
    const mappedField = customFieldForDefinition(childText(definition, "FieldName"), childText(definition, "Alias"));
    if (fieldId && mappedField) fields.set(fieldId, mappedField);
  });

  return fields;
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

function applyCustomFields(
  task: Element,
  item: ProgrammeItem,
  lookup: Map<string, string>,
  customFields: Map<string, keyof ProgrammeItem>,
): ProgrammeItem {
  directChildren(task, "ExtendedAttribute").forEach((attribute) => {
    const fieldId = childText(attribute, "FieldID");
    const fieldName = fieldId ? customFields.get(fieldId) : undefined;
    if (!fieldName) return;
    const raw = resolveLookupValue(childText(attribute, "Value"), childText(attribute, "ValueGUID"), lookup);
    if (fieldName === "roadmapMilestone") item.roadmapMilestone = raw?.toLowerCase() === "yes" || asBool(raw);
    else if (
      fieldName === "discussed" ||
      fieldName === "executiveMilestone" ||
      fieldName === "boardReportable" ||
      fieldName === "decisionRequired" ||
      fieldName === "externalDependency" ||
      fieldName === "dependencyAnchor" ||
      fieldName === "governanceGate"
    ) item[fieldName] = asBool(raw);
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
  const customFields = buildCustomFieldMap(doc);
  const resources = parseResources(doc);
  const assignments = buildAssignments(doc, resources);
  const tasksNode = doc.getElementsByTagName("Tasks")[0];
  if (!tasksNode) throw new Error("No Tasks node was found in this Microsoft Project XML file.");

  const projectStartDate = projectText(doc.documentElement, "StartDate");
  const projectFinishDate = projectText(doc.documentElement, "FinishDate");
  const projectStatusDate = projectText(doc.documentElement, "StatusDate");
  const projectCurrentDate = projectText(doc.documentElement, "CurrentDate");
  const statusDate = projectStatusDate ?? projectCurrentDate;
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
      applyCustomFields(task, item, lookup, customFields);
      item.status = statusFor(item, statusDate);
      return item;
    });

  const normalisedItems = buildSuccessors(buildHierarchy(items));
  return {
    title: projectText(doc.documentElement, "Title") ?? projectText(doc.documentElement, "Name") ?? "Imported programme",
    name: projectText(doc.documentElement, "Name"),
    startDate: projectStartDate,
    finishDate: projectFinishDate,
    statusDate,
    currentDate: projectCurrentDate,
    multipleCriticalPaths: asBool(projectText(doc.documentElement, "MultipleCriticalPaths")),
    items: normalisedItems,
    resources,
    importedAt: new Date().toISOString(),
    sourceFileName,
  };
}
