/**
 * TestHarness.gs
 * One-click verification routines for Assignment Tests 1 through 8.
 * Can be run from the "Ops Automation -> Run Test Scenarios (Tests 1-8)" menu.
 */

function resetDemoData() {
  initializeSystem();
  SpreadsheetApp.getUi().alert('Demo Reset', 'All sheets and dummy tasks have been reset to a clean Cycle 1 state.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * TEST 1 (Recurring Cycle) & TEST 2 (Single-Point Completion)
 * Completes TSK-002 (Portal Task A, 7-day cycle) with a valid Ticket ID and verifies rollover.
 */
function runTest1And2_PortalCompletion() {
  const res = completeTaskById('TSK-002', 'TKT-1042', 'TEST_1_2');
  SpreadsheetApp.getUi().alert(
    'Test 1 & 2 — Portal Task Completion',
    res.message + '\n\nCheck Completion_Log (new row) and Task_Master / Employee_Workspace (rolled to next cycle +7 days).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * TEST 3 (Daily Task Completion — No Ticket ID)
 * Completes TSK-001 (Daily Operations Update) with a WhatsApp update summary.
 */
function runTest3_DailyCompletion() {
  const res = completeTaskById('TSK-001', 'WhatsApp Updated: 45 orders dispatched, 0 pending blockers', 'TEST_3');
  SpreadsheetApp.getUi().alert(
    'Test 3 — Daily Task Completion',
    res.message + '\n\nCheck Completion_Log (proof + timestamp captured) and Task_Master (rolled to next working day).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * TEST 4 (Same-Day Manager Alert)
 * Ensures TSK-002 is Due Today and unalerted, then triggers the Manager Alert engine.
 */
function runTest4_ManagerAlert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(SHEETS.MASTER);
  // Reset TSK-002 (row 3) due date to today & clear alert flag so alert fires cleanly
  master.getRange(3, M_COL.DUE_DATE + 1).setValue(getToday());
  master.getRange(3, M_COL.MANAGER_ALERT_SENT + 1).setValue(false);
  master.getRange(3, M_COL.MANAGER_ALERT_AT + 1).setValue('');
  master.getRange(3, M_COL.MANAGER_ACK + 1).setValue(false);
  master.getRange(3, M_COL.ESCALATED + 1).setValue(false);
  refreshAllViews();

  const res = runManagerAlerts('TEST_4');
  SpreadsheetApp.getUi().alert(
    'Test 4 — Same-Day Manager Alert',
    res.message + '\n\nCheck your Manager email inbox, Manager_Dashboard (Status = Alert Sent), and Automation_Log.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * TEST 5 (Manager Action / Acknowledgement)
 * Records Manager acknowledgement + follow-up note on TSK-002.
 */
function runTest5_ManagerAck() {
  const res = updateManagerAcknowledgement_('TSK-002', true, 'Followed up with Employee B on Slack; ETA 6:30 PM', 'TEST_5');
  SpreadsheetApp.getUi().alert(
    'Test 5 — Manager Acknowledgement',
    res.message + '\n\nCheck Manager_Dashboard (Manager_Ack = TRUE, Status = Manager Acknowledged) and Automation_Log.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * TEST 6 (Leadership Escalation)
 * Sets TSK-003 (Portal Task B) to 2 days overdue, sends Manager Alert (unacknowledged),
 * then triggers Leadership Escalation.
 */
function runTest6_LeadershipEscalation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(SHEETS.MASTER);
  const twoDaysAgo = getToday();
  twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);

  // TSK-003 is row 4 in Task_Master
  master.getRange(4, M_COL.DUE_DATE + 1).setValue(twoDaysAgo);
  master.getRange(4, M_COL.MANAGER_ALERT_SENT + 1).setValue(true);
  master.getRange(4, M_COL.MANAGER_ALERT_AT + 1).setValue(twoDaysAgo);
  master.getRange(4, M_COL.MANAGER_ACK + 1).setValue(false);
  master.getRange(4, M_COL.ESCALATED + 1).setValue(false);
  master.getRange(4, M_COL.ESCALATED_AT + 1).setValue('');
  refreshAllViews();

  const res = runLeadershipEscalations('TEST_6');
  SpreadsheetApp.getUi().alert(
    'Test 6 — Leadership Escalation',
    res.message + '\n\nCheck your Leadership email inbox, Manager_Dashboard (Status = Escalated), and Automation_Log.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * TEST 8 (Idempotency — No Duplicate Alerts, Cycles, or Completions)
 * Re-runs Manager Alert, Leadership Escalation, and a duplicate cycle completion attempt.
 */
function runTest8_IdempotencyCheck() {
  const alertRes = runManagerAlerts('TEST_8_REPEAT');
  const escRes = runLeadershipEscalations('TEST_8_REPEAT');
  const dupComplete = completeTaskById('TSK-002', 'TKT-DUPLICATE', 'TEST_8_REPEAT', 1);

  SpreadsheetApp.getUi().alert(
    'Test 8 — Idempotency Verification',
    `1. Repeat Manager Alert: ${alertRes.message} (Sent: ${alertRes.sentCount})\n` +
    `2. Repeat Leadership Escalation: ${escRes.message} (Escalated: ${escRes.escalatedCount})\n` +
    `3. Duplicate Cycle 1 Completion Attempt: ${dupComplete.message}\n\n` +
    'Verify Automation_Log shows SKIPPED entries and zero duplicate emails/rows were created.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
