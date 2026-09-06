# Savings expectation tracker

A linear stand-in for a savings account. Every pay period the script appends one planned transfer to an Excel ledger. It does not talk to a bank.

Windows Task Scheduler launches the job. The code and conda environment live in WSL Ubuntu. Streamlit can read the same workbook later.

## What it does

On each run the script:

1. Opens `savings_ledger.xlsx` (creates it if missing).
2. Reads `payday_anchor`, `biweekly_amount`, and `starting_balance` from the **Assumptions** sheet.
3. Builds payday dates: the anchor, then every 14 days after that.
4. Appends a `scheduled` row only for paydays **on or before today** that are not already in **Transactions**.
5. Leaves future paydays out of the ledger. Use `--project` or the dashboard to estimate those.

Daily scheduled runs are safe. If today is not a payday, nothing is added. If the PC was off on payday, the next run backfills the missed row.

## Files

| File | Role |
|---|---|
| `savings_tracker.py` | Ledger writer. Run this from Scheduler or by hand. |
| `savings_ledger.xlsx` | Data store. Created on first run. |
| `savings_tracker.log` | Python log next to the script. |
| `app.py` | Optional Streamlit dashboard (read-only). |
| `run_savings_tracker.bat` | Optional wrapper. Prefer calling `wsl.exe` directly from Scheduler. |

## One-time setup

### 1. Edit the script defaults

These values are copied into a **new** workbook only. If `savings_ledger.xlsx` already exists, change the Assumptions sheet instead.

```python
PAYDAY_ANCHOR = date(2026, 9, 11)   # a real payday
BIWEEKLY_AMOUNT = 500.00            # planned transfer each period
STARTING_BALANCE = 0.00             # cash already there before row 1
```

### 2. Environment (WSL + conda)

```bash
cd ~/projects/budget-manager
conda activate budget-manager
pip install openpyxl pandas streamlit
python -c "import openpyxl, sys; print(sys.executable)"
```

You want a path like:

```text
/home/khemra/.conda/envs/budget-manager/bin/python
```

### 3. Create the workbook

From WSL:

```bash
cd ~/projects/budget-manager
/home/khemra/.conda/envs/budget-manager/bin/python savings_tracker.py
```

Or run the Task Scheduler task once after the action is configured.

Close Excel before the script runs. An open workbook will block the save.

## Spreadsheet layout

### Assumptions

| key | meaning |
|---|---|
| `payday_anchor` | First payday. Later paydays are this date + 14, + 28, … |
| `biweekly_amount` | Amount written on **new** scheduled rows. |
| `starting_balance` | Balance before the first Transactions row. Running-balance formulas reference this cell. |

### Transactions

| column | meaning |
|---|---|
| `date` | Payday or the day a manual transfer was logged. |
| `amount_transferred` | Movement that day (negative for a withdrawal). |
| `running_balance` | Excel formula. Do not type over it. |
| `source` | `scheduled` (script) or `manual` (you). |
| `notes` | Optional context. |

After the file exists:

- Change the planned transfer or opening balance on the **Assumptions** tab.
- Do not expect edits to `BIWEEKLY_AMOUNT` / `STARTING_BALANCE` in the `.py` file to affect an existing workbook.
- Extra cash in or out: add a manual row (see CLI). Do not edit `starting_balance` for day-to-day movement.

## Command line

Run from the project directory with the conda env Python:

```bash
python savings_tracker.py
python savings_tracker.py --project 2026-12-31
python savings_tracker.py --add 250 "extra transfer"
python savings_tracker.py --as-of 2026-09-11   # pretend today; useful for tests
```

| Flag | Effect |
|---|---|
| *(none)* | Append any scheduled deposits that are due through today. |
| `--project YYYY-MM-DD` | Print expected balance on that date. Does not write future rows. |
| `--add AMOUNT NOTE` | Append a `manual` row for today. |
| `--as-of YYYY-MM-DD` | Treat that date as today (backfill / testing). |

## Windows Task Scheduler

The project path is Linux. Scheduler must call WSL, not Windows Python and not a `.bat` living on `\\wsl.localhost\...`.

**General**

- Run only when user is logged on
- Run with highest privileges: **off**
- Same Windows account you use interactively

**Trigger**

- Daily, morning (for example 7:00 AM)
- Do not set the trigger to “payday only.” The script decides whether a row is due. Daily also catches a missed payday.

**Action**

| Field | Value |
|---|---|
| Program/script | `C:\Windows\System32\wsl.exe` |
| Add arguments | `-d Ubuntu --cd /home/khemra/projects/budget-manager -- /home/khemra/.conda/envs/budget-manager/bin/python savings_tracker.py` |
| Start in | `C:\Windows\System32` |

Test the same command in Command Prompt before relying on the task. Last Run Result should be `0x0`.

On a non-payday, success still looks like `0 scheduled row(s) added` in `savings_tracker.log`. That is expected.

If Scheduler returns `0x1` and the Python log does not grow, capture output:

```text
-d Ubuntu -- bash -lc "echo START $(date) >> /home/khemra/projects/budget-manager/scheduler_debug.log; cd /home/khemra/projects/budget-manager; /home/khemra/.conda/envs/budget-manager/bin/python savings_tracker.py >> /home/khemra/projects/budget-manager/scheduler_debug.log 2>&1; echo EXIT:$? >> /home/khemra/projects/budget-manager/scheduler_debug.log"
```

## Streamlit dashboard

Read-only view of the same workbook. It recomputes running balance in pandas; it does not evaluate Excel formulas.

```bash
conda activate budget-manager
cd ~/projects/budget-manager
streamlit run app.py
```

Use it for current expected balance, next payday, and “what should this be by date X.” The what-if slider does not write to Excel. Keep writes in the scheduled script (and `--add`).

## Typical calendar

Example: anchor `2026-09-11`, daily task already enabled.

- 9/6–9/10: task runs, 0 rows added, workbook exists.
- 9/11 (or the next successful run after that): one scheduled row.
- 9/25, 10/9, … : next rows.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `ModuleNotFoundError: openpyxl` | Windows Python ran instead of the conda env. Use the full WSL Python path above. |
| Task result `0x1`, no new log line | “Run whether user is logged on or not” or “highest privileges” is on, or the action still points at a UNC `.bat`. |
| Save fails / log shows permission or zip error | `savings_ledger.xlsx` is open in Excel. |
| Wrong payday dates | `payday_anchor` on the Assumptions sheet is still an old date. |
| Amount in new rows is stale | Same: update `biweekly_amount` on Assumptions, not only the `.py` constant. |
| Duplicate payday rows | Unlikely if `source=scheduled` dates are unique. Delete extras by hand if a test `--as-of` was used. |

To rebuild the workbook from the current `.py` defaults: close Excel, delete `savings_ledger.xlsx`, run the script or the task once.


## Useful extra commands already in the starter: 
1. Transfer additional amount 
python savings_tracker.py --add 250 "extra transfer"

2. Prints what the balance would be on that date if planned deposits continue. It does not add future rows to Excel.
Example: Expected balance on 2026-12-31: $4,500.00  (8 future scheduled deposits)
python savings_tracker.py --project 2026-12-31 

--project is the expectation tool: posted activity through today + remaining scheduled deposits until the target date. It does not pre-write 2027 rows into Excel.

3. Pretends today is 2026-10-02.
The script only writes scheduled rows for paydays on or before today. --as-of moves that cutoff so you can:
- Test without waiting for payday
- Backfill if you created the file late and want rows for paydays that already passed
If the anchor is 9/11, --as-of 2026-10-02 would add 9/11, 9/25, and 10/2 (any of those not already in the sheet).
Do not put --as-of on the Task Scheduler action. A wrong date there would invent or skip paydays. For normal use, omit it and let the machine’s real date decide.
python savings_tracker.py --as-of 2026-10-02   # backfill / test



