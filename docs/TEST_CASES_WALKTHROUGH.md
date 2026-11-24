# Test Cases Walkthrough (Tests 1–8)

Every required test case can be verified either **live in the Google Sheet UI** or via the 1-click menu **`Allo Ops Automation` → `Run Test Scenarios (Tests 1-8)`**.

> **Tip for Reviewers:** Keep `Test_Mode = TRUE` in the `Config` tab to test alerts and escalations immediately without waiting for 5:30 PM / 11:00 AM clock triggers.

---

## Verification Matrix

| Test # | Requirement | How to Execute | Expected Result & Where to Verify |
| :--- | :--- | :--- | :--- |
| **Test 1** | **Recurring Cycle Rollover** | Complete `TSK-002` (`Portal Task A`, 7-day cycle) via `Employee_Workspace` (or menu `Test 1 & 2`) | - `Completion_Log` records `TSK-002-C1` with timestamp & Ticket ID.<br>- `Task_Master` & `Employee_Workspace` automatically advance `TSK-002` to **Cycle 2** with `Due_Date = Today + 7 days` (adjusted off non-working days). |
| **Test 2** | **Single-Point Completion** | In `Employee_Workspace`, enter `TKT-1042` in Col I and tick `Mark_Complete` (Col J) | - Employee never leaves `Employee_Workspace`.<br>- Try ticking `Mark_Complete` with an empty Ticket ID first: checkbox auto-unchecks with red error in `Validation_Message`. |
| **Test 3** | **Daily Task Completion (No Ticket ID)** | In `Employee_Workspace`, enter `WhatsApp Updated: 45 orders dispatched` for `TSK-001` and tick `Mark_Complete` (or menu `Test 3`) | - `Completion_Log` captures `Completed_By`, exact timestamp, and WhatsApp summary.<br>- `TSK-001` rolls to **Cycle 2** on the **next working day** (skipping Sunday under `Mon-Sat`). |
| **Test 4** | **Same-Day Manager Alert** | Leave a due-today task incomplete and click `Allo Ops Automation → Run Test Scenarios → Test 4: Same-Day Manager Alert` | - Manager email receives 1 HTML alert email listing the pending task(s).<br>- `Manager_Dashboard` updates `Status = Alert Sent` and increments `Awaiting Manager Ack`.<br>- `Automation_Log` logs `MANAGER_ALERT | SUCCESS`. |
| **Test 5** | **Manager Action & Acknowledgement** | Open `Manager_Dashboard`, tick `Manager_Ack` (Col I) and type a follow-up note in Col J (or run menu `Test 5`) | - `Task_Master` records `Manager_Ack = TRUE`, `Manager_Ack_At` timestamp, and `Manager_Note`.<br>- `Status` changes to `Manager Acknowledged`; `Awaiting Manager Ack` KPI drops.<br>- When employee later completes the task, `Completion_Log` confirms subsequent completion. |
| **Test 6** | **Leadership Escalation** | Click `Allo Ops Automation → Run Test Scenarios → Test 6: Leadership Escalation` (sets `TSK-003` 2 days overdue & unacknowledged) | - Leadership email receives 1 HTML escalation email stating the exact SLA breach reason and manager status.<br>- `Manager_Dashboard` shows `Status = Escalated` and `Escalated = YES`.<br>- `Automation_Log` logs `LEADERSHIP_ESCALATION | SUCCESS`. |
| **Test 7** | **Escalation Timing / SLA Logic** | Inspect `Config` tab & [SLA_AND_DAILY_TASK_DESIGN.md](SLA_AND_DAILY_TASK_DESIGN.md) | - 5:00 PM cutoff → 5:30 PM same-day Manager alert → 2h Manager Ack SLA → 11:00 AM next-working-day Leadership escalation.<br>- All thresholds configurable in `Config` without code edits. |
| **Test 8** | **No Duplicate Alerts, Cycles, or Completions** | Click `Allo Ops Automation → Run Test Scenarios → Test 8: Idempotency / No Duplicates` | - Re-running Manager Alert sends **0** duplicate emails.<br>- Re-running Leadership Escalation sends **0** duplicate emails.<br>- Re-attempting Cycle 1 completion is blocked by `Cycle_Key` check.<br>- `Automation_Log` records `SKIPPED` for all 3 checks. |
