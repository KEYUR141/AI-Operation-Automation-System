/**
 * EscalationEngine.gs
 * Handles:
 *   - Test 4: Same-day Manager alert emails for pending/overdue tasks
 *   - Test 5: Manager acknowledgement & action-note capture from Manager_Dashboard
 *   - Test 6 & 7: Leadership escalation based on SLA timing & unacknowledged alerts
 *   - Test 8: Strict idempotency (no duplicate alert or escalation emails)
 */

/**
 * Combined entry point for time-driven triggers.
 */
function runScheduledChecks(triggerSource) {
  const source = triggerSource || 'SCHEDULED_TRIGGER';
  refreshAllViews();
  runManagerAlerts(source);
  runLeadershipEscalations(source);
}

/**
 * TEST 4: Same-day Manager Alert.
 * Emails the Manager once per pending/overdue cycle when past Manager_Alert_Hour (or in Test_Mode).
 */
function runManagerAlerts(triggerSource) {
  const source = triggerSource || 'MANUAL';
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    logAutomation(source, 'MANAGER_ALERT', 'ALL', '-', '', 'SKIPPED', 'Could not acquire lock.');
    return { sentCount: 0, message: 'System busy.' };
  }

  try {
    const cfg = getConfig();
    if (!cfg.managerEmail) {
      logAutomation(source, 'MANAGER_ALERT', 'ALL', '-', '', 'ERROR', 'Manager_Email is blank in Config.');
      return { sentCount: 0, message: 'Manager_Email is not configured.' };
    }

    const now = new Date();
    const currentHour = now.getHours() + now.getMinutes() / 60;
    if (!cfg.testMode && currentHour < cfg.managerAlertHour) {
      logAutomation(source, 'MANAGER_ALERT', 'ALL', '-', cfg.managerEmail, 'SKIPPED',
        `Current hour (${currentHour.toFixed(2)}) is before Manager_Alert_Hour (${cfg.managerAlertHour}).`);
      return { sentCount: 0, message: 'Before Manager Alert cutoff time.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(SHEETS.MASTER);
    if (!master || master.getLastRow() < 2) return { sentCount: 0, message: 'No tasks found.' };

    const today = getToday();
    const numRows = master.getLastRow() - 1;
    const data = master.getRange(2, 1, numRows, 20).getValues();
    const pendingItems = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[M_COL.TASK_ID] || row[M_COL.ACTIVE] !== true) continue;

      const dueDate = stripTime(row[M_COL.DUE_DATE]);
      const alreadyAlerted = row[M_COL.MANAGER_ALERT_SENT] === true;

      // Trigger if due today or overdue AND alert has not already been sent for this cycle
      if (dueDate <= today && !alreadyAlerted) {
        pendingItems.push({
          rowIndex: i,
          taskId: row[M_COL.TASK_ID],
          taskName: row[M_COL.TASK_NAME],
          assignedTo: row[M_COL.ASSIGNED_TO],
          cycle: row[M_COL.CURRENT_CYCLE],
          dueDate: dueDate,
          daysOverdue: Math.max(0, diffDays(today, dueDate)),
          proofType: row[M_COL.PROOF_TYPE]
        });
      }
    }

    if (pendingItems.length === 0) {
      logAutomation(source, 'MANAGER_ALERT', 'ALL', '-', cfg.managerEmail, 'SKIPPED',
        'No new pending/overdue tasks to alert (duplicates prevented).');
      return { sentCount: 0, message: 'No new pending tasks requiring manager alert.' };
    }

    // Send 1 consolidated email to Manager
    const sheetUrl = ss.getUrl();
    const subject = `[Allo Ops] Same-Day Pending Task Alert (${pendingItems.length} task${pendingItems.length > 1 ? 's' : ''}) — ${formatDate(today)}`;
    const htmlBody = buildManagerAlertHtml_(pendingItems, sheetUrl, cfg);
    const plainBody = pendingItems.map(item =>
      `- ${item.taskId} | ${item.taskName} | ${item.assignedTo} | Cycle ${item.cycle} | Due: ${formatDate(item.dueDate)} (${item.daysOverdue}d overdue)`
    ).join('\n') + `\n\nAction Required: Open Manager_Dashboard and tick Manager_Ack: ${sheetUrl}`;

    MailApp.sendEmail({
      to: cfg.managerEmail,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });

    // Mark Manager_Alert_Sent = true on Task_Master (prevents duplicate alerts)
    for (let j = 0; j < pendingItems.length; j++) {
      const item = pendingItems[j];
      data[item.rowIndex][M_COL.MANAGER_ALERT_SENT] = true;
      data[item.rowIndex][M_COL.MANAGER_ALERT_AT] = now;
      logAutomation(
        source,
        'MANAGER_ALERT',
        item.taskId,
        buildCycleKey(item.taskId, item.cycle),
        cfg.managerEmail,
        'SUCCESS',
        `Same-day manager alert sent for ${item.taskName} (${item.assignedTo}).`
      );
    }

    master.getRange(2, 1, numRows, 20).setValues(data);
    refreshAllViews();

    return { sentCount: pendingItems.length, message: `Sent manager alert for ${pendingItems.length} task(s).` };
  } catch (err) {
    logAutomation(source, 'MANAGER_ALERT', 'ALL', '-', '', 'ERROR', err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * TEST 5: Manager Acknowledgement handler (called by onEdit when Manager_Dashboard is edited).
 */
function handleManagerDashboardEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row < MANAGER_DATA_START_ROW) return;
  if (col !== D_COL.MANAGER_ACK + 1 && col !== D_COL.MANAGER_NOTE + 1) return;

  const rowValues = sheet.getRange(row, 1, 1, 12).getValues()[0];
  const taskId = String(rowValues[D_COL.TASK_ID] || '').trim();
  if (!taskId) return;

  const isAck = rowValues[D_COL.MANAGER_ACK] === true;
  const note = String(rowValues[D_COL.MANAGER_NOTE] || '').trim();

  updateManagerAcknowledgement_(taskId, isAck, note, 'ON_EDIT');
}

/**
 * Programmatic helper to record Manager Acknowledgement (used by onEdit and TestHarness).
 */
function updateManagerAcknowledgement_(taskId, isAck, note, triggerSource) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'System busy.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(SHEETS.MASTER);
    const data = master.getRange(2, 1, master.getLastRow() - 1, 20).getValues();

    let rowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][M_COL.TASK_ID]).trim() === taskId) {
        rowIdx = i;
        break;
      }
    }
    if (rowIdx === -1) return { success: false, message: `Task ${taskId} not found.` };

    const row = data[rowIdx];
    const now = new Date();
    const actor = Session.getActiveUser().getEmail() || 'Manager';

    row[M_COL.MANAGER_ACK] = Boolean(isAck);
    row[M_COL.MANAGER_ACK_AT] = isAck ? (row[M_COL.MANAGER_ACK_AT] || now) : '';
    row[M_COL.MANAGER_NOTE] = note || row[M_COL.MANAGER_NOTE] || '';

    master.getRange(rowIdx + 2, 1, 1, 20).setValues([row]);

    logAutomation(
      triggerSource || 'MANUAL',
      'MANAGER_ACK',
      taskId,
      buildCycleKey(taskId, row[M_COL.CURRENT_CYCLE]),
      actor,
      'SUCCESS',
      `Manager_Ack=${isAck}. Note: "${row[M_COL.MANAGER_NOTE]}"`
    );

    refreshAllViews();
    return { success: true, message: `Updated acknowledgement for ${taskId}.` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * TEST 6 & 7: Leadership Escalation.
 * Escalates alerted tasks to Leadership if:
 *   (A) Manager failed to acknowledge within Manager_Ack_SLA_Hours (or immediately in Test_Mode), OR
 *   (B) Task is still unresolved on the next working day past Leadership_Escalation_Hour (or overdue in Test_Mode).
 */
function runLeadershipEscalations(triggerSource) {
  const source = triggerSource || 'MANUAL';
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    logAutomation(source, 'LEADERSHIP_ESCALATION', 'ALL', '-', '', 'SKIPPED', 'Could not acquire lock.');
    return { escalatedCount: 0, message: 'System busy.' };
  }

  try {
    const cfg = getConfig();
    if (!cfg.leadershipEmail) {
      logAutomation(source, 'LEADERSHIP_ESCALATION', 'ALL', '-', '', 'ERROR', 'Leadership_Email is blank in Config.');
      return { escalatedCount: 0, message: 'Leadership_Email is not configured.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(SHEETS.MASTER);
    if (!master || master.getLastRow() < 2) return { escalatedCount: 0, message: 'No tasks found.' };

    const now = new Date();
    const today = getToday();
    const currentHour = now.getHours() + now.getMinutes() / 60;
    const numRows = master.getLastRow() - 1;
    const data = master.getRange(2, 1, numRows, 20).getValues();
    const escalationItems = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[M_COL.TASK_ID] || row[M_COL.ACTIVE] !== true) continue;

      const alertSent = row[M_COL.MANAGER_ALERT_SENT] === true;
      const alreadyEscalated = row[M_COL.ESCALATED] === true;
      if (!alertSent || alreadyEscalated) continue;

      const dueDate = stripTime(row[M_COL.DUE_DATE]);
      const alertAt = row[M_COL.MANAGER_ALERT_AT] ? new Date(row[M_COL.MANAGER_ALERT_AT]) : now;
      const hoursSinceAlert = (now - alertAt) / (1000 * 60 * 60);
      const managerAck = row[M_COL.MANAGER_ACK] === true;

      let reason = '';
      // Condition A: Manager did not acknowledge within SLA
      if (!managerAck && (cfg.testMode || hoursSinceAlert >= cfg.managerAckSlaHours)) {
        reason = `Manager did not acknowledge alert within ${cfg.managerAckSlaHours}h SLA`;
      }
      // Condition B: Task still unresolved past next-working-day grace period
      else if (dueDate < today && (cfg.testMode || currentHour >= cfg.leadershipEscalationHour)) {
        reason = `Task remained unresolved past next-day ${cfg.leadershipEscalationHour}:00 SLA`;
      }

      if (reason) {
        escalationItems.push({
          rowIndex: i,
          taskId: row[M_COL.TASK_ID],
          taskName: row[M_COL.TASK_NAME],
          assignedTo: row[M_COL.ASSIGNED_TO],
          cycle: row[M_COL.CURRENT_CYCLE],
          dueDate: dueDate,
          daysOverdue: Math.max(0, diffDays(today, dueDate)),
          alertAt: formatDateTime(alertAt),
          managerAck: managerAck ? 'YES' : 'NO',
          managerNote: row[M_COL.MANAGER_NOTE] || '-',
          reason: reason
        });
      }
    }

    if (escalationItems.length === 0) {
      logAutomation(source, 'LEADERSHIP_ESCALATION', 'ALL', '-', cfg.leadershipEmail, 'SKIPPED',
        'No tasks met leadership escalation SLA (or already escalated).');
      return { escalatedCount: 0, message: 'No tasks due for leadership escalation.' };
    }

    const sheetUrl = ss.getUrl();
    const subject = `[ESCALATION] Unresolved Operational Tasks (${escalationItems.length}) — ${formatDate(today)}`;
    const htmlBody = buildEscalationHtml_(escalationItems, sheetUrl);
    const plainBody = escalationItems.map(item =>
      `- ${item.taskId} | ${item.taskName} | ${item.assignedTo} | Due: ${formatDate(item.dueDate)} | Manager Ack: ${item.managerAck} | Reason: ${item.reason}`
    ).join('\n') + `\n\nReview Dashboard: ${sheetUrl}`;

    const mailOptions = {
      to: cfg.leadershipEmail,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    };
    if (cfg.managerEmail && cfg.managerEmail.toLowerCase() !== cfg.leadershipEmail.toLowerCase()) {
      mailOptions.cc = cfg.managerEmail;
    }

    MailApp.sendEmail(mailOptions);

    for (let j = 0; j < escalationItems.length; j++) {
      const item = escalationItems[j];
      data[item.rowIndex][M_COL.ESCALATED] = true;
      data[item.rowIndex][M_COL.ESCALATED_AT] = now;
      logAutomation(
        source,
        'LEADERSHIP_ESCALATION',
        item.taskId,
        buildCycleKey(item.taskId, item.cycle),
        cfg.leadershipEmail,
        'SUCCESS',
        `Escalated ${item.taskName} (${item.assignedTo}). Reason: ${item.reason}`
      );
    }

    master.getRange(2, 1, numRows, 20).setValues(data);
    refreshAllViews();

    return { escalatedCount: escalationItems.length, message: `Escalated ${escalationItems.length} task(s) to Leadership.` };
  } catch (err) {
    logAutomation(source, 'LEADERSHIP_ESCALATION', 'ALL', '-', '', 'ERROR', err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * HTML Email Builders
 */
function buildManagerAlertHtml_(items, sheetUrl, cfg) {
  const rowsHtml = items.map(item => `
    <tr>
      <td style="padding:8px;border:1px solid #ddd;">${item.taskId}</td>
      <td style="padding:8px;border:1px solid #ddd;"><b>${item.taskName}</b></td>
      <td style="padding:8px;border:1px solid #ddd;">${item.assignedTo}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:center;">${item.cycle}</td>
      <td style="padding:8px;border:1px solid #ddd;">${formatDate(item.dueDate)}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:center;">${item.daysOverdue}</td>
    </tr>
  `).join('');

  return `
    <div style="font-family:Arial,sans-serif;color:#202124;">
      <h2 style="color:#b45f06;">Same-Day Pending Recurring Task Alert</h2>
      <p>The following recurring task(s) have not been completed by the cutoff time:</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <thead>
          <tr style="background:#f1f3f4;">
            <th style="padding:8px;border:1px solid #ddd;">Task ID</th>
            <th style="padding:8px;border:1px solid #ddd;">Task Name</th>
            <th style="padding:8px;border:1px solid #ddd;">Assigned To</th>
            <th style="padding:8px;border:1px solid #ddd;">Cycle</th>
            <th style="padding:8px;border:1px solid #ddd;">Due Date</th>
            <th style="padding:8px;border:1px solid #ddd;">Days Overdue</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <p style="margin-top:16px;"><b>Required Manager Action (within ${cfg.managerAckSlaHours} hours):</b></p>
      <ol>
        <li>Open the <a href="${sheetUrl}">Manager_Dashboard</a> tab.</li>
        <li>Follow up with the employee(s) listed above.</li>
        <li>Tick <b>Manager_Ack</b> and enter a brief follow-up note.</li>
      </ol>
      <p style="font-size:12px;color:#5f6368;">Note: If unacknowledged or unresolved past SLA, this will automatically escalate to Leadership.</p>
    </div>
  `;
}

function buildEscalationHtml_(items, sheetUrl) {
  const rowsHtml = items.map(item => `
    <tr>
      <td style="padding:8px;border:1px solid #ddd;">${item.taskId}</td>
      <td style="padding:8px;border:1px solid #ddd;"><b>${item.taskName}</b></td>
      <td style="padding:8px;border:1px solid #ddd;">${item.assignedTo}</td>
      <td style="padding:8px;border:1px solid #ddd;">${formatDate(item.dueDate)}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:center;">${item.managerAck}</td>
      <td style="padding:8px;border:1px solid #ddd;">${item.managerNote}</td>
      <td style="padding:8px;border:1px solid #ddd;color:#b71c1c;"><b>${item.reason}</b></td>
    </tr>
  `).join('');

  return `
    <div style="font-family:Arial,sans-serif;color:#202124;">
      <h2 style="color:#b71c1c;">Leadership Escalation — Unresolved Operational Task(s)</h2>
      <p>The following recurring task(s) breached SLA and require leadership visibility:</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <thead>
          <tr style="background:#fce8e6;">
            <th style="padding:8px;border:1px solid #ddd;">Task ID</th>
            <th style="padding:8px;border:1px solid #ddd;">Task Name</th>
            <th style="padding:8px;border:1px solid #ddd;">Assigned To</th>
            <th style="padding:8px;border:1px solid #ddd;">Due Date</th>
            <th style="padding:8px;border:1px solid #ddd;">Manager Ack</th>
            <th style="padding:8px;border:1px solid #ddd;">Manager Note</th>
            <th style="padding:8px;border:1px solid #ddd;">Escalation Reason</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <p style="margin-top:16px;"><a href="${sheetUrl}">Open Allo Ops Dashboard</a></p>
      <p style="font-size:12px;color:#5f6368;">This escalation is sent once per cycle to prevent inbox spam.</p>
    </div>
  `;
}
