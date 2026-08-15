export type ProgrammeStatus =
  | "complete"
  | "in-progress"
  | "not-started"
  | "late"
  | "at-risk"
  | "blocked"
  | "future";

export type ProgrammeBaseline = {
  number: number;
  start?: string;
  finish?: string;
  duration?: string;
  work?: string;
};

export type ProgrammeDependency = {
  predecessorUid?: string;
  successorUid?: string;
  type?: string;
  lag?: number;
};

export type ProgrammeResource = {
  uid: string;
  name: string;
  initials?: string;
  email?: string;
};

export type ProgrammeItem = {
  uid: string;
  id?: string;
  name: string;
  wbs?: string;
  outlineNumber?: string;
  outlineLevel: number;
  itemType: "summary" | "task" | "milestone";
  isSummary: boolean;
  isMilestone: boolean;
  isCritical: boolean;
  isActive: boolean;
  startDate?: string;
  finishDate?: string;
  duration?: string;
  remainingDuration?: string;
  percentComplete?: number;
  percentWorkComplete?: number;
  earlyStart?: string;
  earlyFinish?: string;
  lateStart?: string;
  lateFinish?: string;
  freeSlack?: number;
  totalSlack?: number;
  startVariance?: number;
  finishVariance?: number;
  baselineStart?: string;
  baselineFinish?: string;
  baselineNumber?: number;
  baselines?: ProgrammeBaseline[];
  predecessors: ProgrammeDependency[];
  successors: ProgrammeDependency[];
  resourceNames?: string[];
  stream?: string;
  roadmapMilestone?: boolean;
  milestoneType?: string;
  approvalBody?: string;
  version?: string;
  visibility?: "Public" | "Internal" | "Restricted" | string;
  roadmapView?: "Governance" | "Delivery" | "Programme" | string;
  milestoneLevel?: string;
  dependencyLevel?: string;
  criticalPathReview?: string;
  ragStatus?: string;
  dateConfidence?: string;
  dateAssumption?: boolean;
  targetMilestone?: string;
  workstreamAccountableOwner?: string;
  deliverySupportRoles?: string;
  projectManagerAssurance?: string;
  discussed?: boolean;
  executiveMilestone?: boolean;
  boardReportable?: boolean;
  decisionRequired?: boolean;
  externalDependency?: boolean;
  dependencyAnchor?: boolean;
  governanceGate?: boolean;
  delayDays?: number;
  delayWorkingDays?: number;
  status: ProgrammeStatus;
  parentUid?: string;
  childUids?: string[];
};

export type ProgrammeSchedule = {
  title: string;
  name?: string;
  startDate?: string;
  finishDate?: string;
  statusDate?: string;
  currentDate?: string;
  multipleCriticalPaths?: boolean;
  items: ProgrammeItem[];
  resources: ProgrammeResource[];
  importedAt: string;
  sourceFileName?: string;
};

export type ProgrammeView =
  | "roadmap"
  | "schedule"
  | "milestones"
  | "governance"
  | "delivery"
  | "release";

export type ProgrammeFilters = {
  stream: string;
  roadmapView: string;
  milestoneType: string;
  approvalBody: string;
  version: string;
  visibility: string;
  status: string;
  criticalOnly: boolean;
  roadmapOnly: boolean;
  showSummaryTasks: boolean;
  delayedOnly: boolean;
  datePreset: "all" | "status-forward" | "current-forward" | "next-30" | "next-60" | "next-90" | "custom";
  dateStart: string;
  dateEnd: string;
  search: string;
};
