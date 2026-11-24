# Daily Task Completion Tracking & Escalation SLA Design

## 1. Daily Task Completion Method (Test 3 — No Ticket ID)

### The Challenge
Daily operational tasks (such as sharing a daily update in a WhatsApp group) do not naturally produce a portal Ticket ID. Asking employees to upload screenshots manually every day creates friction and low adoption; on the other hand, a plain unchecked checkbox with no context invites accidental or fake clicks.

### Proposed Solution (Built into Prototype)
**1-Step Verified Proof inside `Employee_Workspace`:**
1. Employee A posts the daily update in the WhatsApp group as usual.
2. In `Employee_Workspace` (the exact same single-point tab), Employee A enters a **1-line summary of the update** in Column I (e.g., `WhatsApp Updated: 45 orders dispatched, 0 pending blockers`) and ticks **`Mark_Complete`** (Column J).
3. The automation:
   - Validates that the summary is non-empty (`≥ 5` characters).
   - Automatically captures the **logged-in Google account email** (`Session.getActiveUser()`) and **exact server timestamp** (cannot be backdated by the employee).
   - Appends an immutable record to `Completion_Log` and rolls `TSK-001` forward to the **next working day** (skipping configured off-days like Sunday).

### Why This Works Best Practically
* **Zero duplicate work:** Same single-point view as Ticket-based tasks — no switching apps or sheets.
* **Accountability without friction:** Typing a 5-second summary + auto-captured server timestamp prevents blind checkbox clicking while keeping effort minimal.
* **Clear Scale-Up Path:**
  - *Phase 2 (Mobile):* A 2-field Google Form bookmarked on Employee A's phone (`Task` + `1-line WhatsApp summary`) that submits directly into `completeTaskById()`.
  - *Phase 3 (Full Automation):* A WhatsApp Business API / n8n webhook listening for a structured tag (e.g., `#DailyOpsUpdate`) in the WhatsApp group to auto-close `TSK-001` with zero manual sheet entry.

---

## 2. Manager Alerts & Leadership Escalation SLA Logic (Test 7)

### Proposed Timing Schedule

| Milestone | Timing (IST) | System Action | Reasoning |
| :--- | :--- | :--- | :--- |
| **Morning Sync** | **9:00 AM** (Daily) | Refresh `Employee_Workspace` & `Manager_Dashboard` statuses | Employees see fresh `Due Today` items at the start of the workday. |
| **Completion Cutoff** | **5:00 PM** | Expected employee completion deadline | Gives a full working day while leaving time before EOD for recovery. |
| **Same-Day Manager Alert** | **5:30 PM** (Same Day) | Send **1 consolidated HTML email** to Manager for all uncompleted due/overdue tasks; set `Manager_Alert_Sent = TRUE` | A 30-minute grace buffer avoids false alarms on last-minute updates, while still alerting the Manager on the **same day** before close of business. |
| **Manager Acknowledgement SLA** | **Within 2 hours** of alert (or by **10:00 AM** next working day) | Manager opens `Manager_Dashboard`, ticks **`Manager_Ack`**, and enters a brief follow-up note | Proves the Manager actually reviewed the exception and initiated follow-up, rather than ignoring the email. |
| **Leadership Escalation** | **Hourly check** — triggers if:<br>1. `Manager_Ack == FALSE` past **2h SLA**, **OR**<br>2. Task still incomplete past **11:00 AM next working day** | Send **1 consolidated Escalation email** to Leadership (CC Manager) with escalation reason & manager note; set `Escalated = TRUE` | **Reduces Founder dependency:** routine same-day delays are absorbed by the Manager. Leadership is notified *only* on genuine unresolved SLA breaches or missed manager oversight. |

---

### How Repeated / Spam Emails Are Prevented (Test 8)
1. **Per-Cycle State Flags:** Each row in `Task_Master` tracks `Manager_Alert_Sent` and `Escalated`. Once `TRUE`, subsequent trigger runs skip that task for the remainder of the cycle.
2. **Consolidated Digest Emails:** If 3 tasks are overdue at 5:30 PM, the Manager receives **1 clean summary table email** instead of 3 separate emails.
3. **Automatic Reset on Cycle Rollover:** Only when the task is actually completed and rolled to `Current_Cycle + 1` do the flags reset to `FALSE` for the next cycle.
4. **Demo / Reviewer `Test_Mode`:** All timing thresholds live in the `Config` tab. Setting `Test_Mode = TRUE` allows reviewers to trigger alerts and escalations immediately without waiting for 5:30 PM or next day 11:00 AM; setting `Test_Mode = FALSE` enforces strict production clock hours.
