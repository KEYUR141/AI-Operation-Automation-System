/**
 * Config.gs
 * Central constants, column indices, configuration loader, date helpers, and automation logger.
 * Think of this like config.py + utils.py in a Python project.
 */

const SHEETS = {
  CONFIG: 'Config',
  MASTER: 'Task_Master',
  WORKSPACE: 'Employee_Workspace',
  MANAGER: 'Manager_Dashboard',
  COMPLETION_LOG: 'Completion_Log',
  AUTOMATION_LOG: 'Automation_Log'
};

// 0-based column indices for Task_Master (Row 1 = headers, Row 2+ = data)
const M_COL = {
  TASK_ID: 0,
  TASK_NAME: 1,
  ASSIGNED_TO: 2,
  ASSIGNEE_EMAIL: 3,
  FREQUENCY_TYPE: 4,
  CYCLE_DAYS: 5,
  PROOF_TYPE: 6,
  CURRENT_CYCLE: 7,
  DUE_DATE: 8,
  STATUS: 9,
  LAST_COMPLETED_AT: 10,
  LAST_PROOF: 11,
  MANAGER_ALERT_SENT: 12,
  MANAGER_ALERT_AT: 13,
  MANAGER_ACK: 14,
  MANAGER_ACK_AT: 15,
  MANAGER_NOTE: 16,
  ESCALATED: 17,
  ESCALATED_AT: 18,
  ACTIVE: 19
};

// 0-based column indices for Employee_Workspace (Row 3 = headers, Row 4+ = data)
const W_COL = {
  TASK_ID: 0,
  TASK_NAME: 1,
  ASSIGNED_TO: 2,
  FREQUENCY: 3,
  CYCLE_NUM: 4,
  DUE_DATE: 5,
  STATUS: 6,
  PROOF_REQUIRED: 7,
  PROOF_INPUT: 8,
  MARK_COMPLETE: 9,
  VALIDATION_MSG: 10
};

// 0-based column indices for Manager_Dashboard (Row 5 = headers, Row 6+ = data)
const D_COL = {
  TASK_ID: 0,
  TASK_NAME: 1,
  ASSIGNED_TO: 2,
  CYCLE_NUM: 3,
  DUE_DATE: 4,
  DAYS_OVERDUE: 5,
  STATUS: 6,
  MANAGER_ALERT_AT: 7,
  MANAGER_ACK: 8,
  MANAGER_NOTE: 9,
  ESCALATED: 10,
  ESCALATED_AT: 11
};

const WORKSPACE_HEADER_ROW = 3;
const WORKSPACE_DATA_START_ROW = 4;
const WORKSPACE_FILTER_CELL = 'B2';

const MANAGER_HEADER_ROW = 5;
const MANAGER_DATA_START_ROW = 6;

/**
 * Reads key-value pairs from the Config sheet into an object (like a Python dict).
 */
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet) {
    throw new Error('Config sheet not found. Run initializeSystem() first.');
  }

  const values = sheet.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) {
      cfg[key] = values[i][1];
    }
  }

  return {
    managerEmail: String(cfg['Manager_Email'] || '').trim(),
    leadershipEmail: String(cfg['Leadership_Email'] || '').trim(),
    workingDays: String(cfg['Working_Days'] || 'Mon-Sat').trim(),
    taskCutoffHour: Number(cfg['Task_Cutoff_Hour'] || 17),
    managerAlertHour: Number(cfg['Manager_Alert_Hour'] || 17.5),
    managerAckSlaHours: Number(cfg['Manager_Ack_SLA_Hours'] || 2),
    leadershipEscalationHour: Number(cfg['Leadership_Escalation_Hour'] || 11),
    upcomingWindowDays: Number(cfg['Upcoming_Window_Days'] || 30),
    testMode: String(cfg['Test_Mode']).toUpperCase() === 'TRUE' || cfg['Test_Mode'] === true
  };
}

/**
 * Strips time from a Date object (normalizes to 00:00:00 local time).
 */
function stripTime(dateInput) {
  const d = new Date(dateInput);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * Returns today's date normalized to midnight.
 */
function getToday() {
  return stripTime(new Date());
}

/**
 * Checks whether a given date is a working day based on Config ('Mon-Sat' or 'Mon-Fri').
 */
function isWorkingDay(dateInput, workingDaysSetting) {
  const day = new Date(dateInput).getDay(); // 0 = Sun, 6 = Sat
  if (day === 0) return false;
  if (workingDaysSetting === 'Mon-Fri' && day === 6) return false;
  return true;
}

/**
 * Advances a date to the next working day.
 */
function getNextWorkingDay(dateInput, workingDaysSetting) {
  const next = stripTime(dateInput);
  do {
    next.setDate(next.getDate() + 1);
  } while (!isWorkingDay(next, workingDaysSetting));
  return next;
}

/**
 * Calculates the next due date when a cycle is completed.
 * - Daily tasks: next working day after completion/due date (whichever is later, so it never lands in the past).
 * - N-day tasks (7, 15, 30): adds cycleDays to the completed cycle's base date, nudging off non-working days.
 */
function calculateNextDueDate(currentDueDate, cycleDays, frequencyType, workingDaysSetting) {
  const today = getToday();
  const baseDate = stripTime(currentDueDate) > today ? stripTime(currentDueDate) : today;

  if (String(frequencyType).toLowerCase() === 'daily' || Number(cycleDays) <= 1) {
    return getNextWorkingDay(baseDate, workingDaysSetting);
  }

  const next = stripTime(baseDate);
  next.setDate(next.getDate() + Number(cycleDays));
  while (!isWorkingDay(next, workingDaysSetting)) {
    next.setDate(next.getDate() + 1);
  }
  return next;
}

/**
 * Computes whole-day difference: (dateA - dateB) in days.
 */
function diffDays(dateA, dateB) {
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((stripTime(dateA) - stripTime(dateB)) / msPerDay);
}

/**
 * Builds a unique idempotency key for a task cycle (e.g., "TSK-002-C1").
 */
function buildCycleKey(taskId, cycleNumber) {
  return `${taskId}-C${cycleNumber}`;
}

/**
 * Formats a Date for display in emails and logs.
 */
function formatDate(dateInput) {
  if (!dateInput) return '';
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  return Utilities.formatDate(new Date(dateInput), tz, 'yyyy-MM-dd');
}

function formatDateTime(dateInput) {
  if (!dateInput) return '';
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  return Utilities.formatDate(new Date(dateInput), tz, 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Appends an execution/audit event to Automation_Log (monitors successes, skips, and errors).
 */
function logAutomation(triggerSource, action, taskId, cycleKey, recipient, result, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.AUTOMATION_LOG);
    if (!sheet) return;

    const runId = 'RUN-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyyMMdd-HHmmss');
    sheet.appendRow([
      new Date(),
      runId,
      triggerSource || 'MANUAL',
      action || '',
      taskId || '',
      cycleKey || '',
      recipient || '',
      result || 'INFO',
      details || ''
    ]);
  } catch (err) {
    Logger.log('Failed to write Automation_Log: ' + err.message);
  }
}
