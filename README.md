# HSBC SLA Report

Internal tool for The Miles Consultancy. Turns a Fresh CRM monthly export
into the filled HSBC SLA report and sends it to the HSBC contact — from
Bethany's own Microsoft 365 mailbox — with a one-click review and a
per-send audit row.

No database. Config lives in two YAML files the team can edit without a
developer.

---

## First-time setup (Bethany's Windows machine, ~15 minutes)

1. Install **Python 3.11 or newer** from <https://www.python.org/downloads/>.
   Tick the "Add Python to PATH" checkbox in the installer.
2. Unzip `hsbc-sla-tool` somewhere stable, e.g. `C:\Miles\hsbc-sla-tool`.
3. Open the folder in File Explorer and double-click **`run.bat`**. The
   first run creates a virtual environment and installs dependencies.
   That takes a minute or two; subsequent runs are instant.
4. The app opens in the browser at `http://localhost:8501`.
5. Upload March's Fresh export (or whichever month is latest). Click
   **Approve & send** — a small window will show a Microsoft sign-in code.
   Sign in as Bethany once. The token is cached under
   `%USERPROFILE%\.hsbc_sla_tool\token.bin` so you never sign in again
   unless IT forces a credential rotation.

### Before going live (first-month checklist)

These items must be finalised with the team before Andrew at HSBC sees a
report from this tool:

- [ ] `config/settings.yaml` → `email.recipient` set to Andrew's real
      address (confirm with Harrison).
- [ ] `config/settings.yaml` → `email.test_mode: false`.
- [ ] `config/settings.yaml` → `filters.callback_rule.match` confirmed
      with Bethany (see *Setup interview* below).
- [ ] `config/clients.yaml` rules reviewed with Bethany for every
      non-zero line on her last manual report (see *Setup interview*).
- [ ] First run produces numbers that match Bethany's last manual report
      exactly, OR any deltas are understood and signed off.
- [ ] A dry-run send to Bethany's own mailbox (test mode) arrives with
      the xlsx attached and opens cleanly in Excel.

---

## Monthly workflow

1. Export the month's tickets from Fresh as CSV.
2. Double-click `run.bat`.
3. Drop the CSV onto the page.
4. Review the per-client numbers and the *unmapped senders* box.
5. Click **Download for review** to open the filled xlsx and spot-check.
6. Click **Approve & send**.
7. Close the browser when the confirmation screen appears.

The output xlsx is saved to `output/` on disk and attached to the mail —
two copies of the record exist automatically.

---

## Config

Two files in `config/`. Both are plain YAML; any text editor works
(Notepad is fine). Save and re-upload in the app.

### `clients.yaml` — one entry per client row on the HSBC template

```yaml
clients:
  - name: "Honeywell UK"
    rules:
      - field: "Contact ID"
        match: "honeywell\\.uk@fleetlogistics\\.com"
```

- `name` **must match column B of the template exactly**, including
  punctuation. If the template says `NHBC` and the YAML says `Nhbc`, the
  row won't be filled.
- Rules are evaluated top-down; the first matching rule wins. Put the
  most specific clients (e.g. `Honeywell UK`) before the generic ones
  (`Honeywell International`).
- `match` is a case-insensitive regex. `.*` is "anything",
  `\\.` escapes a dot, `\\b` is a word boundary.
- When a ticket matches no rule, it lands in `<UNMAPPED>` and the UI
  surfaces it with a count. Add the sender to `clients.yaml` and
  re-upload — no restart needed.

### `settings.yaml` — email, filters, paths

```yaml
email:
  recipient: "andrew@hsbc.example.com"
  sender_display_name: "Bethany @ Miles Consultancy"
  subject_template: "HSBC SLA Report — {month_label}"
  body_template: |
    Hi Andrew,
    ...
  test_mode: true
  test_recipient: "bethany@milesconsultancy.co.uk"

filters:
  exclude_groups: ["Updates"]
  callback_rule:
    field: "Subject"
    match: ".*call\\s*back.*"
```

- `{month_label}` in the subject/body is replaced with e.g. `March 2026`.
- `test_mode: true` redirects mail to `test_recipient` and shows a banner
  in the UI. Flip to `false` only after the dry-run checklist is done.

---

## Setup interview — do this with Bethany before first live send

Schedule 30 minutes. Open her March 2026 report and the March Fresh
export side by side.

1. Walk through every non-zero row on her Services table. For each:
   - Which CSV rows did you count? What identified them as that client?
2. The CSV has ~9,000 rows; the Services total is ~945. What filters
   drop the other ~8,000?
   - `Updates` group = 3,775 of them (already handled).
   - Confirm anything else.
3. The **Call backs** section — which tickets count? Likely a Group
   value or a Subject pattern. Set `filters.callback_rule` accordingly.
4. The **Reporting Mailbox / Payrolls / Reconciliation** section — is
   that in this CSV, a different export, or manual? For v1 it's out of
   scope; Bethany will keep filling it by hand.
5. Edge cases she mentioned (e.g. Ashfield) — codify them in `clients.yaml`.

Every answer should land in one of the two YAML files. Those files are
her knowledge, written down.

---

## How to add a new client

1. Open `config/clients.yaml`.
2. Add an entry above `- name: "YOUFIBRE LIMITED"` (alphabetical is
   polite but not required):
   ```yaml
     - name: "Acme Widgets Ltd"
       rules:
         - field: "Full name"
           match: ".*acme.*"
         - field: "Contact ID"
           match: ".*@acmewidgets\\."
   ```
3. Add a matching row in column B of `templates/HSBC_SLA_template.xlsx`
   (inside the Services section between the header and the blank row
   before the Call backs section — and also in the Call backs table if
   they'll have callbacks).
4. Re-upload the CSV in the app.

---

## Audit log

`audit_log.csv` in the project root. One row per successful send.
Append-only: the app never rewrites or deletes rows. Fields:

| column | meaning |
|---|---|
| `timestamp_utc` | when the send completed (UTC ISO-8601) |
| `user_email` | whose mailbox the mail went from |
| `source_csv_sha256` | hash of the Fresh export |
| `output_xlsx_sha256` | hash of the filled report |
| `month_label` | e.g. `March 2026` |
| `services_total_within`, `services_total_outside` | |
| `callbacks_total_within`, `callbacks_total_outside` | |
| `unmapped_count` | total tickets that didn't match any client |
| `recipient` | inbox the mail was sent to |
| `graph_message_id` | Microsoft Graph request id for traceability |

Hand this file to Harrison when compliance asks "did you definitely send
the right thing that month?".

---

## If something goes wrong

**"That doesn't look like a Fresh export."** — the CSV is missing one
of `Resolution status`, `Group`, `Created time`, `Full name`,
`Contact ID`, or `Subject`. Fresh may have renamed a column. Re-export
with default columns.

**The sign-in screen doesn't appear.** — close the browser tab and try
again; the device-code flow will print a code in the page. If the code
expires, close the tab and retry.

**Send fails with a Graph error.** — the error card keeps you on the
review screen; try again. If the error mentions `Mail.Send`, IT needs
to grant Bethany the `Mail.Send` scope (it's a default permission for
most M365 accounts but some regulated tenants restrict it).

**Unmapped senders you expected to be mapped.** — open
`config/clients.yaml` and widen the regex. Test by clicking the file
uploader again and re-dropping the same CSV.

---

## For developers

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python3 -m pytest
.venv/bin/streamlit run app.py
```

The processor is pure functions — everything testable without touching
disk, a network, or a clock. The mailer is isolated behind a single
`send_mail(token, ...)` call; swap it for a Gmail equivalent if Miles
moves off Microsoft 365 and nothing else changes.

### Swapping to Google Workspace

Replace `mailer.py` with a Gmail API implementation that exposes the
same `acquire_token` / `send_mail` / `sign_out` contract. `app.py`
calls only those three functions; no other code changes.
