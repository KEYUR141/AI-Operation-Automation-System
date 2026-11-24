/**
 * TriggersAndMenu.gs
 * Adds the custom "Allo Ops Automation" menu to Google Sheets,
 * routes live spreadsheet edits (Employee_Workspace & Manager_Dashboard),
 * and installs all required event + time-driven triggers idempotently.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const testMenu = ui.createMenu('Run Test Scenarios (Tests 1-8)')
    .addItem('Reset Demo Data (Clean Slate)', 'resetDemoData')
    .addSeparator()
    .addItem('Test 1 & 2: Complete 7-Day Portal Task', 'runTest1And2_PortalCompletion')
    .addItem('Test 3: Complete Daily WhatsApp Task', 'runTest3_DailyCompletion')
    .addItem('Test 4: Same-Day Manager Alert', 'runTest4_ManagerAlert')
    .addItem('Test 5: Manager Acknowledgement', 'runTest5_ManagerAck')
    .addItem('Test 6: Leadership Escalation', 'runTest6_LeadershipEscalation')
    .addItem('Test 8: Idempotency / No Duplicates', 'runTest8_IdempotencyCheck');

  ui.createMenu('Allo Ops Automation')
    .addItem('Refresh All Views', 'refreshAllViews')
    .addSeparator()
    .addItem('Run Manager Alert Check Now', 'menuRunManagerAlerts')
    .addItem('Run Leadership Escalation Check Now', 'menuRunLeadershipEscalations')
    .addSeparator()
    .addSubMenu(testMenu)
    .addSeparator()
    .addItem('Install Automatic Triggers (onEdit + Schedule)', 'installTriggers')
    .addItem('Re-Initialize Entire System', 'initializeSystem')
    .addToUi();
}

/**
 * Installable onEdit trigger handler.
 * Runs with owner authorization so cross-sheet sync & user tracking work for all editors.
 */
function handleSheetEdit(e) {
  if (!e || !e.range) return;
  try {
    const sheetName = e.range.getSheet().getName();
    if (sheetName === SHEETS.WORKSPACE) {
      handleEmployeeWorkspaceEdit(e);
    } else if (sheetName === SHEETS.MANAGER) {
      handleManagerDashboardEdit(e);
    }
  } catch (err) {
    logAutomation('ON_EDIT', 'HANDLE_EDIT', '', '', '', 'ERROR', err.message);
  }
}

/**
 * Installs all required triggers cleanly (deletes old triggers first so there are never duplicates).
 */
function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existing.length; i++) {
    ScriptApp.deleteTrigger(existing[i]);
  }

  // 1. Live single-point edit handler (Employee_Workspace + Manager_Dashboard)
  ScriptApp.newTrigger('handleSheetEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // 2. Daily morning view refresh (9:00 AM)
  ScriptApp.newTrigger('refreshAllViews')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // 3. Same-day Manager alert check (5:30 PM)
  ScriptApp.newTrigger('runManagerAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(30)
    .create();

  // 4. Hourly Leadership escalation check (only emails when SLA conditions are met)
  ScriptApp.newTrigger('runLeadershipEscalations')
    .timeBased()
    .everyHours(1)
    .create();

  logAutomation('SETUP', 'INSTALL_TRIGGERS', 'ALL', '-', Session.getEffectiveUser().getEmail(), 'SUCCESS',
    'Installed 4 triggers: handleSheetEdit (onEdit), refreshAllViews (9 AM), runManagerAlerts (5:30 PM), runLeadershipEscalations (hourly).');

  SpreadsheetApp.getUi().alert(
    'Triggers Installed',
    '4 automatic triggers are now active:\n' +
    '1. Live onEdit handler (Employee_Workspace & Manager_Dashboard)\n' +
    '2. Daily 9:00 AM view refresh\n' +
    '3. Daily 5:30 PM same-day Manager alert check\n' +
    '4. Hourly Leadership escalation SLA check',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Menu wrappers with friendly UI notifications.
 */
function menuRunManagerAlerts() {
  const res = runManagerAlerts('MENU_MANUAL');
  SpreadsheetApp.getUi().alert('Manager Alert Check', res.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRunLeadershipEscalations() {
  const res = runLeadershipEscalations('MENU_MANUAL');
  SpreadsheetApp.getUi().alert('Leadership Escalation Check', res.message, SpreadsheetApp.getUi().ButtonSet.OK);
}
