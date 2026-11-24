/**
 * WorkspaceSync.gs
 * Keeps Task_Master statuses fresh and syncs the two user-facing views:
 *   1. Employee_Workspace (single-point view for employees)
 *   2. Manager_Dashboard  (team-wide pending/overdue/escalation view for manager)
 */

function refreshAllViews() {
  recalculateTaskStatuses_();
  refreshEmployeeWorkspace();
  refreshManagerDashboard();
  SpreadsheetApp.flush();
}

/**
 * Recalculates Status on Task_Master based on Due_Date and alert/ack/escalation flags.
 */
function recalculateTaskStatuses_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(SHEETS.MASTER);
  if (!master || master.getLastRow() < 2) return;

  const today = getToday();
  const numRows = master.getLastRow() - 1;
  const data = master.getRange(2, 1, numRows, 20).getValues();
  let changed = false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[M_COL.TASK_ID] || row[M_COL.ACTIVE] !== true) continue;

    const dueDate = stripTime(row[M_COL.DUE_DATE]);
    const alertSent = row[M_COL.MANAGER_ALERT_SENT] === true;
    const managerAck = row[M_COL.MANAGER_ACK] === true;
    const escalated = row[M_COL.ESCALATED] === true;

    let newStatus = 'Upcoming';
    if (escalated) {
      newStatus = 'Escalated';
    } else if (managerAck) {
      newStatus = 'Manager Acknowledged';
    } else if (alertSent) {
      newStatus = 'Alert Sent';
    } else if (dueDate < today) {
      newStatus = 'Overdue';
    } else if (dueDate.getTime() === today.getTime()) {
      newStatus = 'Due Today';
    }

    if (row[M_COL.STATUS] !== newStatus) {
      data[i][M_COL.STATUS] = newStatus;
      changed = true;
    }
  }

  if (changed) {
    master.getRange(2, 1, numRows, 20).setValues(data);
  }
}

/**
 * Populates Employee_Workspace from Task_Master, respecting the employee filter in B2.
 */
function refreshEmployeeWorkspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(SHEETS.MASTER);
  const ws = ss.getSheetByName(SHEETS.WORKSPACE);
  if (!master || !ws) return;

  const cfg = getConfig();
  const today = getToday();
  const filterValue = String(ws.getRange(WORKSPACE_FILTER_CELL).getValue() || 'All Employees').trim();

  // Clear existing data rows
  const maxRows = Math.max(ws.getLastRow() - WORKSPACE_DATA_START_ROW + 1, 20);
  const clearRange = ws.getRange(WORKSPACE_DATA_START_ROW, 1, maxRows, 11);
  clearRange.clearContent().clearDataValidations().setBackground(null).setFontColor('#000000');

  if (master.getLastRow() < 2) return;

  const masterData = master.getRange(2, 1, master.getLastRow() - 1, 20).getValues();
  const rowsToShow = [];

  for (let i = 0; i < masterData.length; i++) {
    const r = masterData[i];
    if (!r[M_COL.TASK_ID] || r[M_COL.ACTIVE] !== true) continue;

    const assignedTo = String(r[M_COL.ASSIGNED_TO]).trim();
    if (filterValue !== 'All Employees' && assignedTo !== filterValue) continue;

    const dueDate = stripTime(r[M_COL.DUE_DATE]);
    const daysUntilDue = diffDays(dueDate, today);

    // Show overdue, due today, and tasks within the configured upcoming window
    if (daysUntilDue <= cfg.upcomingWindowDays) {
      rowsToShow.push([
        r[M_COL.TASK_ID],
        r[M_COL.TASK_NAME],
        assignedTo,
        r[M_COL.FREQUENCY_TYPE],
        r[M_COL.CURRENT_CYCLE],
        dueDate,
        r[M_COL.STATUS],
        r[M_COL.PROOF_TYPE],
        '',    // Enter_Ticket_ID_or_Proof (fresh input cell)
        false, // Mark_Complete checkbox
        ''     // Validation_Message
      ]);
    }
  }

  // Sort by Due_Date ascending (most urgent / overdue at the top)
  rowsToShow.sort((a, b) => a[W_COL.DUE_DATE] - b[W_COL.DUE_DATE]);

  if (rowsToShow.length > 0) {
    ws.getRange(WORKSPACE_DATA_START_ROW, 1, rowsToShow.length, 11).setValues(rowsToShow);
    ws.getRange(WORKSPACE_DATA_START_ROW, W_COL.DUE_DATE + 1, rowsToShow.length, 1).setNumberFormat('yyyy-mm-dd');
    ws.getRange(WORKSPACE_DATA_START_ROW, W_COL.PROOF_INPUT + 1, rowsToShow.length, 1).setBackground('#fffde7');
    ws.getRange(WORKSPACE_DATA_START_ROW, W_COL.MARK_COMPLETE + 1, rowsToShow.length, 1).insertCheckboxes();
  }
}

/**
 * Populates Manager_Dashboard with all active tasks currently Due Today, Overdue, Alerted, or Escalated.
 */
function refreshManagerDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(SHEETS.MASTER);
  const dash = ss.getSheetByName(SHEETS.MANAGER);
  if (!master || !dash) return;

  const today = getToday();

  // Clear existing data rows
  const maxRows = Math.max(dash.getLastRow() - MANAGER_DATA_START_ROW + 1, 20);
  const clearRange = dash.getRange(MANAGER_DATA_START_ROW, 1, maxRows, 12);
  clearRange.clearContent().clearDataValidations().setBackground(null);

  if (master.getLastRow() < 2) return;

  const masterData = master.getRange(2, 1, master.getLastRow() - 1, 20).getValues();
  const rowsToShow = [];

  for (let i = 0; i < masterData.length; i++) {
    const r = masterData[i];
    if (!r[M_COL.TASK_ID] || r[M_COL.ACTIVE] !== true) continue;

    const status = String(r[M_COL.STATUS]);
    const needsManagerView = ['Due Today', 'Overdue', 'Alert Sent', 'Manager Acknowledged', 'Escalated'].includes(status);
    if (!needsManagerView) continue;

    const dueDate = stripTime(r[M_COL.DUE_DATE]);
    const daysOverdue = Math.max(0, diffDays(today, dueDate));

    rowsToShow.push([
      r[M_COL.TASK_ID],
      r[M_COL.TASK_NAME],
      r[M_COL.ASSIGNED_TO],
      r[M_COL.CURRENT_CYCLE],
      dueDate,
      daysOverdue,
      status,
      r[M_COL.MANAGER_ALERT_AT] || '',
      r[M_COL.MANAGER_ACK] === true,
      r[M_COL.MANAGER_NOTE] || '',
      r[M_COL.ESCALATED] === true ? 'YES' : 'NO',
      r[M_COL.ESCALATED_AT] || ''
    ]);
  }

  // Sort: Escalated / most overdue at the top
  rowsToShow.sort((a, b) => b[D_COL.DAYS_OVERDUE] - a[D_COL.DAYS_OVERDUE]);

  if (rowsToShow.length > 0) {
    dash.getRange(MANAGER_DATA_START_ROW, 1, rowsToShow.length, 12).setValues(rowsToShow);
    dash.getRange(MANAGER_DATA_START_ROW, D_COL.DUE_DATE + 1, rowsToShow.length, 1).setNumberFormat('yyyy-mm-dd');
    dash.getRange(MANAGER_DATA_START_ROW, D_COL.MANAGER_ALERT_AT + 1, rowsToShow.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    dash.getRange(MANAGER_DATA_START_ROW, D_COL.ESCALATED_AT + 1, rowsToShow.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    dash.getRange(MANAGER_DATA_START_ROW, D_COL.MANAGER_ACK + 1, rowsToShow.length, 1).insertCheckboxes();
    dash.getRange(MANAGER_DATA_START_ROW, D_COL.MANAGER_NOTE + 1, rowsToShow.length, 1).setBackground('#e8f0fe');
  }
}
