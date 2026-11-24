/**
 * Setup.gs
 * 1-click initializer that creates all 6 sheets, formatting, KPI formulas,
 * data validations, conditional formatting, and the 4 required dummy tasks.
 */

function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getEffectiveUser().getEmail() || 'manager@example.com';

  setupConfigSheet_(ss, userEmail);
  setupTaskMasterSheet_(ss, userEmail);
  setupEmployeeWorkspaceSheet_(ss);
  setupManagerDashboardSheet_(ss);
  setupCompletionLogSheet_(ss);
  setupAutomationLogSheet_(ss);

  // Remove default blank "Sheet1" if it still exists and we have our 6 tabs
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  logAutomation('SETUP', 'INITIALIZE_SYSTEM', 'ALL', '-', userEmail, 'SUCCESS', 'All 6 sheets created with dummy tasks and formulas.');

  // If WorkspaceSync.gs is already added, populate the views immediately
  if (typeof refreshAllViews === 'function') {
    refreshAllViews();
  }

  SpreadsheetApp.flush();
}

/**
 * Helper: gets an existing sheet or creates it, and resets its content/formatting.
 */
function getOrResetSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }
  return sheet;
}

/**
 * 1. Config Sheet
 */
function setupConfigSheet_(ss, defaultEmail) {
  const sheet = getOrResetSheet_(ss, SHEETS.CONFIG);
  const rows = [
    ['Setting_Key', 'Value', 'Description'],
    ['Manager_Email', defaultEmail, 'Email that receives same-day pending task alerts (e.g., you+manager@gmail.com)'],
    ['Leadership_Email', defaultEmail, 'Email that receives leadership escalations (e.g., you+leadership@gmail.com)'],
    ['Working_Days', 'Mon-Sat', 'Working week: Mon-Sat (Sunday off) or Mon-Fri (Sat-Sun off)'],
    ['Task_Cutoff_Hour', 17, 'Hour (24h) by which due tasks should be completed (17 = 5:00 PM)'],
    ['Manager_Alert_Hour', 17.5, 'Hour (24h) when same-day Manager alert fires (17.5 = 5:30 PM)'],
    ['Manager_Ack_SLA_Hours', 2, 'Hours allowed for Manager to acknowledge alert before escalation'],
    ['Leadership_Escalation_Hour', 11, 'Hour (24h) on next working day when unresolved tasks escalate (11 = 11:00 AM)'],
    ['Upcoming_Window_Days', 30, 'Days ahead to display upcoming tasks in Employee_Workspace'],
    ['Test_Mode', true, 'TRUE = ignore clock-hour checks so reviewers can trigger alerts/escalations immediately']
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  styleHeaderRow_(sheet, 1, 3, '#1a73e8');

  // Dropdown for Working_Days and checkbox for Test_Mode
  const workingDaysRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Mon-Sat', 'Mon-Fri'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B4').setDataValidation(workingDaysRule);
  sheet.getRange('B10').insertCheckboxes();

  sheet.getRange('B2:B10').setBackground('#fff8e1');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(3, 520);
}

/**
 * 2. Task_Master Sheet (Source of Truth)
 */
function setupTaskMasterSheet_(ss, defaultEmail) {
  const sheet = getOrResetSheet_(ss, SHEETS.MASTER);
  const headers = [
    'Task_ID', 'Task_Name', 'Assigned_To', 'Assignee_Email',
    'Frequency_Type', 'Cycle_Days', 'Proof_Type', 'Current_Cycle',
    'Due_Date', 'Status', 'Last_Completed_At', 'Last_Proof',
    'Manager_Alert_Sent', 'Manager_Alert_At', 'Manager_Ack', 'Manager_Ack_At',
    'Manager_Note', 'Escalated', 'Escalated_At', 'Active'
  ];

  const today = getToday();
  const plus2 = new Date(today); plus2.setDate(today.getDate() + 2);
  const plus5 = new Date(today); plus5.setDate(today.getDate() + 5);

  const dummyTasks = [
    ['TSK-001', 'Daily Operations Update', 'Employee A', defaultEmail, 'Daily', 1, 'WhatsApp Update Summary', 1, today, 'Due Today', '', '', false, '', false, '', '', false, '', true],
    ['TSK-002', 'Portal Task A', 'Employee B', defaultEmail, 'Every 7 Days', 7, 'Ticket ID', 1, today, 'Due Today', '', '', false, '', false, '', '', false, '', true],
    ['TSK-003', 'Portal Task B', 'Employee C', defaultEmail, 'Every 15 Days', 15, 'Ticket ID', 1, plus2, 'Upcoming', '', '', false, '', false, '', '', false, '', true],
    ['TSK-004', 'Portal Task C', 'Employee D', defaultEmail, 'Every 30 Days', 30, 'Ticket ID', 1, plus5, 'Upcoming', '', '', false, '', false, '', '', false, '', true]
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, dummyTasks.length, headers.length).setValues(dummyTasks);

  styleHeaderRow_(sheet, 1, headers.length, '#202124');
  sheet.getRange(2, M_COL.DUE_DATE + 1, 20, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, M_COL.LAST_COMPLETED_AT + 1, 20, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, M_COL.MANAGER_ALERT_AT + 1, 20, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, M_COL.MANAGER_ACK_AT + 1, 20, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, M_COL.ESCALATED_AT + 1, 20, 1).setNumberFormat('yyyy-mm-dd hh:mm');

  // Checkboxes for boolean flags
  sheet.getRange(2, M_COL.MANAGER_ALERT_SENT + 1, dummyTasks.length, 1).insertCheckboxes();
  sheet.getRange(2, M_COL.MANAGER_ACK + 1, dummyTasks.length, 1).insertCheckboxes();
  sheet.getRange(2, M_COL.ESCALATED + 1, dummyTasks.length, 1).insertCheckboxes();
  sheet.getRange(2, M_COL.ACTIVE + 1, dummyTasks.length, 1).insertCheckboxes();

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 3. Employee_Workspace Sheet (Single-Point Update View)
 */
function setupEmployeeWorkspaceSheet_(ss) {
  const sheet = getOrResetSheet_(ss, SHEETS.WORKSPACE);

  sheet.getRange('A1:K1').merge()
    .setValue('OPS AUTOMATION — EMPLOYEE PENDING TASKS (SINGLE-POINT UPDATE)')
    .setBackground('#0f4c81').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);

  sheet.getRange('A2').setValue('Filter Employee:').setFontWeight('bold');
  sheet.getRange(WORKSPACE_FILTER_CELL).setValue('All Employees').setBackground('#e8f0fe').setFontWeight('bold');
  const filterRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['All Employees', 'Employee A', 'Employee B', 'Employee C', 'Employee D'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(WORKSPACE_FILTER_CELL).setDataValidation(filterRule);

  sheet.getRange('C2:K2').merge()
    .setValue('Instructions: Enter your Ticket ID (e.g., TKT-1042) OR Daily WhatsApp update summary in Column I, then tick Mark_Complete in Column J. Cycle rollover happens automatically.')
    .setFontStyle('italic').setFontColor('#5f6368');

  const headers = [
    'Task_ID', 'Task_Name', 'Assigned_To', 'Frequency', 'Cycle #',
    'Due_Date', 'Status', 'Proof_Required', 'Enter_Ticket_ID_or_Proof',
    'Mark_Complete', 'Validation_Message'
  ];
  sheet.getRange(WORKSPACE_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, WORKSPACE_HEADER_ROW, headers.length, '#1a73e8');

  // Conditional formatting for Status column (Col G)
  applyStatusConditionalFormatting_(sheet, 'G4:G50');

  sheet.setFrozenRows(WORKSPACE_HEADER_ROW);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(W_COL.PROOF_INPUT + 1, 240);
  sheet.setColumnWidth(W_COL.VALIDATION_MSG + 1, 260);
}

/**
 * 4. Manager_Dashboard Sheet (Team Visibility + Manager Acknowledgement)
 */
function setupManagerDashboardSheet_(ss) {
  const sheet = getOrResetSheet_(ss, SHEETS.MANAGER);

  sheet.getRange('A1:L1').merge()
    .setValue('OPS AUTOMATION — MANAGER MONITORING & ESCALATION DASHBOARD')
    .setBackground('#b71c1c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);

  // KPI Labels (Row 2) and Live Formulas (Row 3)
  const kpiLabels = [
    'Total Active Tasks', 'Due Today', 'Pending / Overdue', 'Awaiting Manager Ack', 'Escalated to Leadership'
  ];
  const kpiFormulas = [
    `=COUNTIF(${SHEETS.MASTER}!T2:T, TRUE)`,
    `=COUNTIFS(${SHEETS.MASTER}!T2:T, TRUE, ${SHEETS.MASTER}!J2:J, "Due Today")`,
    `=COUNTIFS(${SHEETS.MASTER}!T2:T, TRUE, ${SHEETS.MASTER}!J2:J, "Overdue") + COUNTIFS(${SHEETS.MASTER}!T2:T, TRUE, ${SHEETS.MASTER}!J2:J, "Alert Sent") + COUNTIFS(${SHEETS.MASTER}!T2:T, TRUE, ${SHEETS.MASTER}!J2:J, "Manager Acknowledged")`,
    `=COUNTIFS(${SHEETS.MASTER}!T2:T, TRUE, ${SHEETS.MASTER}!M2:M, TRUE, ${SHEETS.MASTER}!O2:O, FALSE)`,
    `=COUNTIFS(${SHEETS.MASTER}!T2:T, TRUE, ${SHEETS.MASTER}!R2:R, TRUE)`
  ];

  for (let i = 0; i < kpiLabels.length; i++) {
    const col = i * 2 + 1;
    sheet.getRange(2, col, 1, 2).merge().setValue(kpiLabels[i])
      .setBackground('#f1f3f4').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(3, col, 1, 2).merge().setFormula(kpiFormulas[i])
      .setBackground('#ffffff').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.getRange('A4:L4').merge()
    .setValue('Manager Action: Review overdue/alerted tasks below. Tick Manager_Ack (Col I) and add an Action Note (Col J) to record acknowledgement.')
    .setFontStyle('italic').setFontColor('#5f6368');

  const headers = [
    'Task_ID', 'Task_Name', 'Assigned_To', 'Cycle #', 'Due_Date',
    'Days_Overdue', 'Status', 'Manager_Alert_At', 'Manager_Ack',
    'Manager_Action_Note', 'Escalated', 'Escalated_At'
  ];
  sheet.getRange(MANAGER_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, MANAGER_HEADER_ROW, headers.length, '#37474f');

  applyStatusConditionalFormatting_(sheet, 'G6:G50');

  sheet.setFrozenRows(MANAGER_HEADER_ROW);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(D_COL.MANAGER_NOTE + 1, 250);
}

/**
 * 5. Completion_Log Sheet (Immutable Audit Trail)
 */
function setupCompletionLogSheet_(ss) {
  const sheet = getOrResetSheet_(ss, SHEETS.COMPLETION_LOG);
  const headers = [
    'Log_ID', 'Cycle_Key', 'Task_ID', 'Task_Name', 'Cycle_Number',
    'Assigned_To', 'Completed_By', 'Due_Date', 'Completed_At',
    'SLA_Status', 'Proof_Type', 'Proof_Value'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length, '#137333');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 6. Automation_Log Sheet (Monitoring & Failure Log)
 */
function setupAutomationLogSheet_(ss) {
  const sheet = getOrResetSheet_(ss, SHEETS.AUTOMATION_LOG);
  const headers = [
    'Timestamp', 'Run_ID', 'Trigger_Source', 'Action',
    'Task_ID', 'Cycle_Key', 'Recipient', 'Result', 'Details'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, 1, headers.length, '#4a148c');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(9, 420);
}

/**
 * Shared styling helpers
 */
function styleHeaderRow_(sheet, row, numCols, bgHex) {
  sheet.getRange(row, 1, 1, numCols)
    .setBackground(bgHex)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function applyStatusConditionalFormatting_(sheet, rangeA1) {
  const range = sheet.getRange(rangeA1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Escalated').setBackground('#f4cccc').setFontColor('#990000').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Alert Sent').setBackground('#fce5cd').setFontColor('#b45f06').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Overdue').setBackground('#f8d7da').setFontColor('#842029').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Due Today').setBackground('#fff3cd').setFontColor('#664d03').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Manager Acknowledged').setBackground('#cfe2ff').setFontColor('#084298').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Upcoming').setBackground('#d1e7dd').setFontColor('#0f5132').setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}
