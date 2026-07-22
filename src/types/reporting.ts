export type ReportingPeriod = {
  start?: Date;
  end?: Date;
  label: string;
};

export type WeeklySummary = {
  id: string;
  weekEnding?: string;
  meetingDate?: string;
  overallRag?: string;
  ragRationale?: string;
  topicsDiscussed?: string;
  keyProgress?: string;
  whatChanged?: string;
  keyRisksOrIssues?: string;
  decisionsMade?: string;
  decisionsNeeded?: string;
  priorityActions?: string;
  openingLine?: string;
  askSteerNeeded?: string;
  mainBlocker?: string;
  goLiveConfidence?: string;
  ragMovement?: string;
  steerRequired?: string;
  updateType?: string;
  lastUpdated?: string;
};

export type TrackerRisk = {
  id: string;
  dateRaised?: string;
  dashboardFlag?: boolean;
  lastDiscussedDate?: string;
  stream?: string;
  title: string;
  statement?: string;
  cause?: string;
  impact?: string;
  likelihood?: string;
  impactRating?: string;
  rag?: string;
  mitigation?: string;
  owner?: string;
  targetDate?: string;
  status?: string;
  updateType?: string;
  latestUpdate?: string;
};

export type TrackerIssue = {
  id: string;
  dateRaised?: string;
  dashboardFlag?: boolean;
  lastDiscussedDate?: string;
  stream?: string;
  title: string;
  statement?: string;
  impact?: string;
  priority?: string;
  rag?: string;
  requiredAction?: string;
  requiredDecision?: string;
  owner?: string;
  targetDate?: string;
  status?: string;
  updateType?: string;
  latestUpdate?: string;
};

export type TrackerAction = {
  id: string;
  meetingDate?: string;
  dashboardFlag?: boolean;
  stream?: string;
  title: string;
  description?: string;
  status?: string;
  owner?: string;
  priority?: string;
  dueDate?: string;
  updateType?: string;
  latestUpdate?: string;
};

export type TrackerDecision = {
  id: string;
  decisionDate?: string;
  decisionRequiredBy?: string;
  decisionRequiredByLabel?: string;
  dashboardFlag?: boolean;
  decisionType?: string;
  lastDiscussedDate?: string;
  stream?: string;
  title: string;
  statement?: string;
  decisionMaker?: string;
  owner?: string;
  status?: string;
  updateType?: string;
  latestUpdate?: string;
};

export type TrackerChange = {
  id: string;
  dateRaised?: string;
  lastDiscussedDate?: string;
  dashboardFlag?: boolean;
  changeType?: string;
  stream?: string;
  title: string;
  description?: string;
  reason?: string;
  impactOnTime?: string;
  impactOnScope?: string;
  impactOnCost?: string;
  impactOnQualityOrBenefits?: string;
  decisionRequired?: string;
  decisionMaker?: string;
  owner?: string;
  status?: string;
  updateType?: string;
  latestUpdate?: string;
};

export type MeetingMinute = {
  id: string;
  meetingDate?: string;
  meetingName?: string;
  stream?: string;
  topic?: string;
  summary?: string;
  outcomeType?: string;
  updateType?: string;
  latestUpdate?: string;
};

export type TrackerData = {
  sourceFileName?: string;
  importedAt: string;
  weeklySummaries: WeeklySummary[];
  risks: TrackerRisk[];
  issues: TrackerIssue[];
  actions: TrackerAction[];
  decisions: TrackerDecision[];
  changes: TrackerChange[];
  minutes: MeetingMinute[];
};

export type ReleasePlanData = {
  sourceFileName?: string;
  importedAt: string;
  versions: unknown[];
  scopeItems: unknown[];
  readinessItems: unknown[];
};
