/**
 * TaskEngine.gs
 * Handles single-point task completion from Employee_Workspace,
 * proof validation (Ticket ID & Daily WhatsApp Update), idempotent logging,
 * and automatic recurring cycle rollover.
 */

/**
 * Entry point called by onEdit(e) when Employee_Workspace is edited.
 */
function handleEmployeeWorkspaceEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Case 1: Employee filter dropdown changed in B2
  if (e.range.getA1Notation() === WORKSPACE_FILTER_CELL) {
    refreshEmployeeWorkspace();
    return;
  }

  // Case 2: Mark_Complete checkbox ticked in Column J (col index 10)
  if (row >= WORKSPACE_DATA_START_ROW && col === W_COL.MARK_COMPLETE + 1 && e.value === 'TRUE') {
    processWorkspaceRowCompletion_(sheet, row, 'ON_EDIT');
  }
}

/**
 * Validates and processes a single row completion in Employee_Workspace.
 */
function processWorkspaceRowCompletion_(wsSheet, rowNum, triggerSource) {
  const rowValues = wsSheet.getRange(rowNum, 1, 1, 11).getValues()[0];
  const taskId = String(rowValues[W_COL.TASK_ID] || '').trim();
  const cycleNum = Number(rowValues[W_COL.CYCLE_NUM]);
  const proofRequired = String(rowValues[W_COL.PROOF_REQUIRED] || '').trim();
  const proofInput = String(rowValues[W_COL.PROOF_INPUT] || '').trim();

  if (!taskId) return;

  // Validate proof before touching Task_Master
  const validation = validateProof_(proofRequired, proofInput);
  if (!validation.valid) {
    wsSheet.getRange(rowNum, W_COL.MARK_COMPLETE + 1).setValue(false);
    wsSheet.getRange(rowNum, W_COL.VALIDATION_MSG + 1)
      .setValue(validation.message)
      .setFontColor('#b71c1c');
    logAutomation(triggerSource, 'COMPLETE_TASK', taskId, buildCycleKey(taskId, cycleNum), '', 'REJECTED', validation.message);
    return;
  }

  // Clear any old error and run core completion engine
  wsSheet.getRange(rowNum, W_COL.VALIDATION_MSG + 1).setValue('Processing...').setFontColor('#1a73e8');
  const result = completeTaskById(taskId, proofInput, triggerSource, cycleNum);

  if (!result.success) {
    wsSheet.getRange(rowNum, W_COL.MARK_COMPLETE + 1).setValue(false);
    wsSheet.getRange(rowNum, W_COL.VALIDATION_MSG + 1)
      .setValue(result.message)
      .setFontColor('#b71c1c');
  }
}

/**
 * Proof validation rules (prevents empty or junk completions).
 */
function validateProof_(proofType, proofValue) {
  if (!proofValue) {
    return {
      valid: false,
      message: proofType === 'Ticket ID'
        ? 'Ticket ID required (e.g., TKT-1042) before marking complete.'
        : 'Enter WhatsApp update summary (min 5 chars) before marking complete.'
    };
  }

  if (proofType === 'Ticket ID') {
    const ticketRegex = /^[A-Za-z0-9][A-Za-z0-9\-_\/#]{3,39}$/;
    if (!ticketRegex.test(proofValue)) {
      return {
        valid: false,
        message: 'Invalid Ticket ID format. Use 4-40 alphanumeric chars (e.g., TKT-1042).'
      };
    }
  } else {
    if (proofValue.length < 5) {
      return {
        valid: false,
        message: 'Daily proof too short. Add a brief update summary (min 5 chars).'
      };
    }
  }

  return { valid: true, message: 'OK' };
}

/**
 * Core completion engine (used by both onEdit and TestHarness).
 * Uses LockService + Cycle_Key check for strict idempotency (Test 8).
 */
function completeTaskById(taskId, proofValue, triggerSource, expectedCycle) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'System busy. Please try again in a few seconds.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(SHEETS.MASTER);
    const logSheet = ss.getSheetByName(SHEETS.COMPLETION_LOG);
    const cfg = getConfig();

    const masterData = master.getRange(2, 1, master.getLastRow() - 1, 20).getValues();
    let rowIdx = -1;
    for (let i = 0; i < masterData.length; i++) {
      if (String(masterData[i][M_COL.TASK_ID]).trim() === taskId) {
        rowIdx = i;
        break;
      }
    }

    if (rowIdx === -1) {
      return { success: false, message: `Task ${taskId} not found in Task_Master.` };
    }

    const row = masterData[rowIdx];
    const currentCycle = Number(row[M_COL.CURRENT_CYCLE]);
    const proofType = String(row[M_COL.PROOF_TYPE]);

    const validation = validateProof_(proofType, String(proofValue || '').trim());
    if (!validation.valid) {
      return { success: false, message: validation.message };
    }

    // Guard against stale UI cycle mismatch
    if (expectedCycle && Number(expectedCycle) !== currentCycle) {
      const msg = `Cycle ${expectedCycle} already closed (current is Cycle ${currentCycle}).`;
      logAutomation(triggerSource || 'MANUAL', 'COMPLETE_TASK', taskId, buildCycleKey(taskId, expectedCycle), '', 'SKIPPED', msg);
      refreshAllViews();
      return { success: false, message: msg };
    }

    const cycleKey = buildCycleKey(taskId, currentCycle);

    // Idempotency check against Completion_Log (Test 8)
    if (isCycleAlreadyLogged_(logSheet, cycleKey)) {
      const msg = `Duplicate completion prevented: ${cycleKey} already exists in Completion_Log.`;
      logAutomation(triggerSource || 'MANUAL', 'COMPLETE_TASK', taskId, cycleKey, '', 'SKIPPED', msg);
      return { success: false, message: msg };
    }

    const now = new Date();
    const today = getToday();
    const dueDate = stripTime(row[M_COL.DUE_DATE]);
    const slaStatus = today <= dueDate ? 'On Time' : 'Late';
    const completedBy = Session.getActiveUser().getEmail() || row[M_COL.ASSIGNED_TO];

    // 1. Append immutable record to Completion_Log
    const logId = 'LOG-' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyyMMdd-HHmmss');
    logSheet.appendRow([
      logId,
      cycleKey,
      taskId,
      row[M_COL.TASK_NAME],
      currentCycle,
      row[M_COL.ASSIGNED_TO],
      completedBy,
      dueDate,
      now,
      slaStatus,
      proofType,
      String(proofValue).trim()
    ]);

    // 2. Roll Task_Master forward to next cycle and reset alert/escalation flags
    const nextCycle = currentCycle + 1;
    const nextDueDate = calculateNextDueDate(dueDate, row[M_COL.CYCLE_DAYS], row[M_COL.FREQUENCY_TYPE], cfg.workingDays);

    row[M_COL.CURRENT_CYCLE] = nextCycle;
    row[M_COL.DUE_DATE] = nextDueDate;
    row[M_COL.LAST_COMPLETED_AT] = now;
    row[M_COL.LAST_PROOF] = String(proofValue).trim();
    row[M_COL.MANAGER_ALERT_SENT] = false;
    row[M_COL.MANAGER_ALERT_AT] = '';
    row[M_COL.MANAGER_ACK] = false;
    row[M_COL.MANAGER_ACK_AT] = '';
    row[M_COL.MANAGER_NOTE] = '';
    row[M_COL.ESCALATED] = false;
    row[M_COL.ESCALATED_AT] = '';

    master.getRange(rowIdx + 2, 1, 1, 20).setValues([row]);

    logAutomation(
      triggerSource || 'MANUAL',
      'COMPLETE_TASK',
      taskId,
      cycleKey,
      completedBy,
      'SUCCESS',
      `Closed ${cycleKey} (${slaStatus}, Proof: ${proofValue}). Rolled to Cycle ${nextCycle}, Next Due: ${formatDate(nextDueDate)}.`
    );

    // 3. Refresh Employee_Workspace & Manager_Dashboard
    refreshAllViews();

    return {
      success: true,
      message: `Completed ${cycleKey}. Next cycle (${nextCycle}) due ${formatDate(nextDueDate)}.`,
      nextCycle: nextCycle,
      nextDueDate: formatDate(nextDueDate)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Checks Completion_Log for an existing Cycle_Key to guarantee no duplicate completion entries.
 */
function isCycleAlreadyLogged_(logSheet, cycleKey) {
  if (!logSheet || logSheet.getLastRow() < 2) return false;
  const keys = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === cycleKey) {
      return true;
    }
  }
  return false;
}
