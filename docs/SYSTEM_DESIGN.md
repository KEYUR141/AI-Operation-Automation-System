# System Design & Operational Guardrails

## 1. End-to-End Process Map
`Input → Process → Decision → Output → Exception → Escalation`

| Stage | Description |
| :--- | :--- |
| **Input** | Recurring task definitions in `Task_Master` (4 dummy tasks: Daily, 7-day, 15-day, 30-day) + employee proof inputs (`Ticket ID` or `WhatsApp Update Summary`) in `Employee_Workspace` + manager acknowledgement in `Manager_Dashboard`. |
| **Process** | 1. Daily 9:00 AM refresh computes live statuses (`Upcoming`, `Due Today`, `Overdue`, `Alert Sent`, `Manager Acknowledged`, `Escalated`).<br>2. Live `onEdit` trigger validates employee proof and executes cycle rollover inside `LockService`.<br>3. Scheduled triggers check SLA deadlines at 5:30 PM (Manager Alert) and hourly (Leadership Escalation). |
| **Decision** | - Is proof valid? (Regex for Ticket ID; min-length check for Daily WhatsApp summary)<br>- Is `Cycle_Key` already in `Completion_Log`? (Idempotency check)<br>- Is task due/overdue and `Manager_Alert_Sent == FALSE`?<br>- Did Manager acknowledge within SLA, and is task still unresolved past next-day 11:00 AM? |
| **Output** | - Immutable row appended to `Completion_Log` (who, when, SLA status, proof)<br>- `Task_Master` rolled forward to `Current_Cycle + 1` and next working-day `Due_Date`<br>- Refreshed `Employee_Workspace` & `Manager_Dashboard` KPIs<br>- Consolidated HTML emails to Manager / Leadership + audit entry in `Automation_Log` |
| **Exception** | - Blank or malformed Ticket ID → Checkbox auto-unchecked, red error message displayed in `Validation_Message`, logged as `REJECTED`.<br>- Duplicate trigger / double-click → Blocked by `LockService` + `Cycle_Key` / boolean flags, logged as `SKIPPED`.<br>- Script/runtime error → Caught in `try/catch`, written to `Automation_Log` with `Result = ERROR`. |
| **Escalation** | Employee misses 5:00 PM cutoff → **Manager alerted at 5:30 PM same day** → Manager must acknowledge within 2 hours → **Leadership escalated ONLY if** Manager fails to acknowledge OR task stays unresolved past 11:00 AM next working day. |

---

## 2. How the System Prevents Incorrect Changes (Guardrails)

1. **Single-Point Editing with Separation of Concerns:**
   - Employees only interact with `Employee_Workspace` (and only two columns: `Enter_Ticket_ID_or_Proof` and `Mark_Complete`).
   - `Task_Master`, `Completion_Log`, and `Automation_Log` are system-managed tabs (can be sheet-protected to Owner/Script-only in production so no employee can overwrite due dates or delete history).
2. **Strict Input Validation Before State Changes:**
   - Portal tasks require a valid Ticket ID (`4–40` alphanumeric/hyphen chars, e.g., `TKT-1042`).
   - Daily tasks require a non-empty update summary (`≥ 5` characters).
   - Invalid submissions are rejected *before* touching `Task_Master` or `Completion_Log`.
3. **Concurrency Control (`LockService`):**
   - All completion, acknowledgement, alert, and escalation functions acquire `LockService.getDocumentLock()` to prevent race conditions if two triggers or edits fire simultaneously.
4. **Strict Idempotency (`Cycle_Key` + State Flags):**
   - Every completed cycle writes a unique `Cycle_Key` (`TSK-002-C1`) to `Completion_Log`. Re-running completion on an already-closed cycle is rejected with `SKIPPED`.
   - Every alert and escalation checks `Manager_Alert_Sent` and `Escalated` flags before sending email, guaranteeing **zero duplicate emails**.
5. **Auditability & Failure Monitoring (`Automation_Log`):**
   - Every run logs `Timestamp | Run_ID | Trigger_Source | Action | Task_ID | Cycle_Key | Recipient | Result | Details`, providing full observability into successes, skips, rejections, and failures.

---

## 3. What Should & Should Not Be Automated

| Automate | Keep Human-in-the-Loop |
| :--- | :--- |
| Next due-date & working-day calculations | Actually executing the portal work / posting the WhatsApp operational update |
| Pending task detection & employee workspace filtering | Manager's judgement call & follow-up action note when an employee is delayed |
| Same-day Manager alerts & SLA-based Leadership escalations | Leadership intervention on repeated or high-risk operational bottlenecks |
| Audit logging, KPI aggregation & duplicate prevention | Modifying task frequencies or disabling active tasks in `Task_Master` |
