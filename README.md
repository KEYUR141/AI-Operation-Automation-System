# Allo Innoware — Recurring Task Tracking, Pending Reports & Escalation Automation

A working **Google Sheets + Google Apps Script** operational automation system built for the **Allo Innoware — AI Operations / Founder's Office** assignment.

It automates the full recurring operational task lifecycle:
`Task Becomes Due` → `Employee Single-Point Action` → `Proof Validation & Cycle Rollover` → `Same-Day Manager Alert` → `Manager Acknowledgement` → `Leadership Escalation` — with strict idempotency (zero duplicate cycles, logs, or emails).

---

## System Architecture

```text
Task_Master (Source of Truth)
  ├──► Employee_Workspace (Single-Point UI + Employee Filter + Ticket/Daily Proof)
  │         │  [Live installable onEdit + LockService]
  │         ├── Invalid proof ──► Uncheck + Red Validation Message + Log REJECTED
  │         └── Valid proof   ──► Append Completion_Log (unique Cycle_Key)
  │                               Increment Cycle # + Calculate Next Working-Day Due Date
  │                               Reset alert/ack/escalation flags & refresh views
  │
  └──► Scheduled Triggers (Cron)
            ├── 5:30 PM Same Day ──► Email Manager (once/cycle) + set Manager_Alert_Sent = TRUE
            ├── Manager_Dashboard ──► Manager ticks Manager_Ack + adds Action Note
            └── Hourly SLA Check  ──► Email Leadership (once/cycle) ONLY if unacknowledged past SLA
                                      OR unresolved past next-working-day 11:00 AM
```

---

## Google Sheet Tabs (6 Tabs)

| Tab | Purpose |
| :--- | :--- |
| **`Config`** | Manager & Leadership emails, working days (`Mon-Sat` / `Mon-Fri`), SLA hours, and `Test_Mode` toggle |
| **`Task_Master`** | Central task database (frequency, cycle #, due date, status, alert/ack/escalation flags) |
| **`Employee_Workspace`** | Single-point view where employees filter their tasks, enter Ticket ID / Daily Proof, and tick `Mark_Complete` |
| **`Manager_Dashboard`** | Live KPI header (`COUNTIF`/`COUNTIFS` formulas) + team-wide overdue/alerted tasks + Manager Acknowledgement checkboxes |
| **`Completion_Log`** | Immutable append-only audit trail keyed by unique `Cycle_Key` (e.g., `TSK-002-C1`) |
| **`Automation_Log`** | Execution, email, and error monitor (`SUCCESS` / `SKIPPED` / `REJECTED` / `ERROR`) |

---

## Repository Structure

```text
AI-Operation-Automation-System/
├── README.md
├── apps-script/
│   ├── Config.gs               # Constants, column indices, config loader, working-day date math, logger
│   ├── Setup.gs                # 1-click initializer (creates all 6 tabs, formulas, formatting & 4 dummy tasks)
│   ├── WorkspaceSync.gs        # Status recalculation + sync for Employee_Workspace & Manager_Dashboard
│   ├── TaskEngine.gs           # Single-point completion, proof validation, idempotency, cycle rollover
│   ├── EscalationEngine.gs     # Same-day Manager alerts, Manager acknowledgement, Leadership escalation
│   ├── TriggersAndMenu.gs      # Custom "Allo Ops Automation" menu, installable onEdit, time-driven triggers
│   └── TestHarness.gs          # 1-click verification runners for Assignment Tests 1–8
└── docs/
    ├── SYSTEM_DESIGN.md            # Process map, guardrails against incorrect changes & failure monitoring
    ├── SLA_AND_DAILY_TASK_DESIGN.md # Daily task tracking method & Escalation SLA timing rationale
    └── TEST_CASES_WALKTHROUGH.md   # Step-by-step verification guide for all 8 required tests
```

---

## Quick Setup (5 Minutes)

1. Open a new Google Sheet (`sheets.new`) → **Extensions → Apps Script**.
2. Add the 7 `.gs` files from `apps-script/` into the Apps Script editor and save (`Ctrl + S`).
3. Run **`initializeSystem`** from `Setup.gs` and approve permissions.
4. Reload the Google Sheet tab → click **`Allo Ops Automation` → `Install Automatic Triggers (onEdit + Schedule)`**.
5. Update `Manager_Email` and `Leadership_Email` in the **`Config`** tab (optional; defaults to your Google account).
6. Use **`Allo Ops Automation` → `Run Test Scenarios (Tests 1-8)`** or interact directly with `Employee_Workspace` and `Manager_Dashboard` to verify every test case.

---

## Documentation
* [System Design & Data Guardrails](docs/SYSTEM_DESIGN.md)
* [Daily Task Completion & Escalation SLA Design](docs/SLA_AND_DAILY_TASK_DESIGN.md)
* [Test Cases 1–8 Walkthrough](docs/TEST_CASES_WALKTHROUGH.md)
