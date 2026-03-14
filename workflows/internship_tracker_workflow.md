# Internship Application Tracker — Agentic Workflow

> **Claude Code workflow** — run via `claude` CLI. Scans Gmail for internship/spring week applications, builds a live `.xlsx` tracker, and syncs successful applications to Google Calendar.

---

## Architecture Overview

```
Gmail (MCP) → Claude Parser → Excel Tracker (.xlsx) → Google Calendar (MCP)
                    ↑
             Anthropic API
             (claude-haiku-4-5 for extraction,
              claude-sonnet-4-6 for ambiguous cases)
```

**Pipeline stages:**
1. `fetch` — Pull emails from Gmail matching application keywords
2. `parse` — Extract structured application data from each email thread
3. `deduplicate` — Merge threads for the same role into one record
4. `write` — Create/update `applications.xlsx` with all records
5. `sync_calendar` — For offers/interviews confirmed, push Google Calendar events

---

## Prerequisites

```bash
pip install anthropic openpyxl pandas python-dateutil google-auth google-auth-oauthlib google-api-python-client
```

Set your Anthropic API key:
```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

---

## Configuration

Edit `config.py` to match your preferences:

```python
# config.py

# --- Search settings ---
SEARCH_QUERY = """
(subject:("application received") OR subject:("thank you for applying") OR
 subject:("spring week") OR subject:("summer internship") OR subject:("off-cycle") OR
 subject:("online assessment") OR subject:("video interview") OR subject:("hackerrank") OR
 subject:("offer") OR subject:("unfortunately") OR subject:("application unsuccessful") OR
 subject:("next steps") OR from:(noreply@greenhouse.io) OR from:(noreply@lever.co) OR
 from:(donotreply@workday.com))
"""
DATE_CUTOFF = "2024/09/01"          # Only emails after this date (YYYY/MM/DD)
MAX_THREADS = 300                    # Cap on email threads to fetch

# --- Output ---
OUTPUT_XLSX = "applications.xlsx"
SHEET_NAME = "Applications"

# --- Calendar ---
CALENDAR_ID = "primary"             # or your calendar email, e.g. ryan@gmail.com
TIMEZONE = "Europe/London"

# --- Status taxonomy ---
STATUSES = [
    "Applied",
    "Online Assessment",
    "First Round Interview",
    "Second Round Interview",
    "Assessment Centre",
    "Offer",
    "Rejected",
    "Withdrawn",
    "Pending",
]

# --- Role types to track ---
ROLE_TYPES = [
    "Spring Week",
    "Summer Internship",
    "Off-Cycle Internship",
    "Graduate Scheme",
    "Part-Time / Other",
]
```

---

## Spreadsheet Schema

The output `applications.xlsx` has the following columns:

| Column | Type | Description |
|---|---|---|
| `company` | string | Company name |
| `role` | string | Role/programme title |
| `role_type` | enum | Spring Week / Summer Internship / etc. |
| `division` | string | Division/desk (e.g. IBD, S&T, Tech) |
| `location` | string | Office location |
| `date_applied` | date | Date of initial application |
| `deadline` | date | Application deadline (if found) |
| `status` | enum | Current status (see taxonomy above) |
| `last_updated` | date | Date of most recent status email |
| `next_step` | string | e.g. "Complete HireVue by 20 Mar" |
| `next_deadline` | date | Deadline for next action item |
| `notes` | string | Free text / extracted context |
| `email_thread_id` | string | Gmail thread ID for traceability |
| `calendar_event_id` | string | GCal event ID (populated on sync) |

---

## Main Script

```python
# tracker.py
import json, os
from datetime import datetime
from anthropic import Anthropic
from config import *
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

client = Anthropic()

# ─────────────────────────────────────────
# STEP 1: FETCH EMAILS VIA GMAIL MCP
# ─────────────────────────────────────────

def fetch_application_emails():
    """Use Claude + Gmail MCP to pull all relevant email threads."""
    print("📥 Fetching application emails from Gmail...")

    response = client.beta.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4096,
        tools=[],
        mcp_servers=[{"type": "url", "url": "https://gmail.mcp.claude.com/mcp", "name": "gmail"}],
        messages=[{
            "role": "user",
            "content": f"""Search Gmail for internship/spring week application emails.
Use the search query: {SEARCH_QUERY} after:{DATE_CUTOFF}
Fetch up to {MAX_THREADS} threads. For each thread, retrieve the full content.
Return a JSON array where each element has:
  - thread_id: string
  - subject: string  
  - from: string
  - date: ISO date string
  - body_snippet: first 500 chars of most recent message in thread
  - all_subjects: array of all subjects in thread (to track status progression)
Return ONLY valid JSON, no markdown."""
        }],
        betas=["mcp-client-2025-04-04"]
    )

    raw = "".join(b.text for b in response.content if b.type == "text")
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Fallback: extract JSON from response
        import re
        match = re.search(r'\[.*\]', raw, re.DOTALL)
        return json.loads(match.group()) if match else []


# ─────────────────────────────────────────
# STEP 2: PARSE EMAILS INTO STRUCTURED DATA
# ─────────────────────────────────────────

PARSE_SYSTEM = """You are an expert at parsing internship application emails.
Given email thread data, extract structured application information.
Return ONLY a JSON object with these exact keys:
  company, role, role_type, division, location,
  date_applied (ISO date), deadline (ISO date or null),
  status (from: Applied/Online Assessment/First Round Interview/
          Second Round Interview/Assessment Centre/Offer/Rejected/Withdrawn/Pending),
  last_updated (ISO date), next_step (string or null),
  next_deadline (ISO date or null), notes

Be conservative: if unsure, use "Pending" for status and null for unknowns.
role_type must be one of: Spring Week/Summer Internship/Off-Cycle Internship/Graduate Scheme/Part-Time / Other"""

def parse_email_thread(thread: dict) -> dict:
    """Parse a single email thread into a structured application record."""
    model = "claude-haiku-4-5-20251001"
    
    prompt = f"""Parse this email thread into a structured application record:

Subject: {thread.get('subject')}
From: {thread.get('from')}
Date: {thread.get('date')}
Thread subjects (chronological): {thread.get('all_subjects', [])}
Body snippet: {thread.get('body_snippet', '')}
Thread ID: {thread.get('thread_id')}"""

    response = client.messages.create(
        model=model,
        max_tokens=512,
        system=PARSE_SYSTEM,
        messages=[{"role": "user", "content": prompt}]
    )
    
    raw = response.content[0].text
    try:
        record = json.loads(raw)
    except json.JSONDecodeError:
        # Escalate ambiguous cases to Sonnet
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            system=PARSE_SYSTEM,
            messages=[{"role": "user", "content": prompt}]
        )
        record = json.loads(response.content[0].text)
    
    record["email_thread_id"] = thread.get("thread_id", "")
    record["calendar_event_id"] = ""
    return record


# ─────────────────────────────────────────
# STEP 3: DEDUPLICATE RECORDS
# ─────────────────────────────────────────

def deduplicate(records: list[dict]) -> list[dict]:
    """Merge multiple records for the same company+role, keeping latest status."""
    seen = {}
    for r in records:
        key = (r.get("company", "").lower().strip(), r.get("role", "").lower().strip())
        if key not in seen:
            seen[key] = r
        else:
            # Keep the record with the more advanced status
            existing = seen[key]
            if STATUSES.index(r.get("status", "Pending")) > STATUSES.index(existing.get("status", "Pending")):
                r["date_applied"] = existing.get("date_applied")  # preserve original apply date
                seen[key] = r
    return list(seen.values())


# ─────────────────────────────────────────
# STEP 4: WRITE XLSX
# ─────────────────────────────────────────

STATUS_COLOURS = {
    "Applied":                  "D6E4F0",
    "Online Assessment":        "FFF3CD",
    "First Round Interview":    "D4EDDA",
    "Second Round Interview":   "A8D5BA",
    "Assessment Centre":        "7EC8A4",
    "Offer":                    "28A745",
    "Rejected":                 "F8D7DA",
    "Withdrawn":                "E2E3E5",
    "Pending":                  "FFFFFF",
}

HEADERS = [
    "Company", "Role", "Role Type", "Division", "Location",
    "Date Applied", "Deadline", "Status", "Last Updated",
    "Next Step", "Next Deadline", "Notes", "Thread ID", "Calendar Event ID"
]

RECORD_KEYS = [
    "company", "role", "role_type", "division", "location",
    "date_applied", "deadline", "status", "last_updated",
    "next_step", "next_deadline", "notes", "email_thread_id", "calendar_event_id"
]

def write_xlsx(records: list[dict], path: str = OUTPUT_XLSX):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    # Header row
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # Data rows
    for row, rec in enumerate(records, 2):
        status = rec.get("status", "Pending")
        row_colour = STATUS_COLOURS.get(status, "FFFFFF")
        fill = PatternFill("solid", fgColor=row_colour)

        for col, key in enumerate(RECORD_KEYS, 1):
            val = rec.get(key, "")
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = fill
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    # Column widths
    col_widths = [22, 28, 18, 18, 14, 14, 14, 20, 14, 35, 14, 35, 20, 22]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Stats summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Status"
    ws2["B1"] = "Count"
    for i, s in enumerate(STATUSES, 2):
        ws2.cell(row=i, column=1, value=s)
        ws2.cell(row=i, column=2,
                 value=f'=COUNTIF({SHEET_NAME}!H:H,"{s}")')
    ws2["A1"].font = Font(bold=True)
    ws2["B1"].font = Font(bold=True)
    ws2["A12"] = "Total Applications"
    ws2["B12"] = f'=COUNTA({SHEET_NAME}!A:A)-1'

    wb.save(path)
    print(f"✅ Saved tracker → {path}  ({len(records)} records)")


# ─────────────────────────────────────────
# STEP 5: SYNC TO GOOGLE CALENDAR
# ─────────────────────────────────────────

CALENDAR_TRIGGER_STATUSES = {
    "First Round Interview", "Second Round Interview",
    "Assessment Centre", "Offer"
}

def sync_calendar_events(records: list[dict]) -> list[dict]:
    """Create Google Calendar events for interviews and offers."""
    print("📅 Syncing calendar events...")
    
    events_to_create = [
        r for r in records
        if r.get("status") in CALENDAR_TRIGGER_STATUSES
        and not r.get("calendar_event_id")
        and r.get("next_deadline")
    ]

    if not events_to_create:
        print("   No new calendar events needed.")
        return records

    event_specs = json.dumps([{
        "title": f"{r['company']} — {r['status']}",
        "date": r["next_deadline"],
        "description": f"Role: {r['role']}\nNext step: {r.get('next_step', '')}\nNotes: {r.get('notes', '')}",
        "record_index": i
    } for i, r in enumerate(records) if r in events_to_create])

    response = client.beta.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2048,
        mcp_servers=[{"type": "url", "url": "https://gcal.mcp.claude.com/mcp", "name": "gcal"}],
        messages=[{
            "role": "user",
            "content": f"""Create Google Calendar events for these internship milestones.
Calendar: {CALENDAR_ID}, Timezone: {TIMEZONE}
Set duration to 1 hour unless it's an assessment centre (set to 4 hours).
For each event created, return its event ID.

Events: {event_specs}

Return a JSON array of objects: {{record_index, event_id}}"""
        }],
        betas=["mcp-client-2025-04-04"]
    )

    raw = "".join(b.text for b in response.content if b.type == "text")
    try:
        import re
        match = re.search(r'\[.*\]', raw, re.DOTALL)
        event_ids = json.loads(match.group()) if match else []
        
        # Write event IDs back to records
        event_map = {e["record_index"]: e["event_id"] for e in event_ids}
        for idx, r in enumerate(records):
            if idx in event_map:
                r["calendar_event_id"] = event_map[idx]
    except Exception as e:
        print(f"   ⚠️  Calendar sync partial error: {e}")

    return records


# ─────────────────────────────────────────
# ENTRYPOINT
# ─────────────────────────────────────────

def run():
    print("🚀 Starting internship tracker pipeline\n")
    
    # 1. Fetch
    threads = fetch_application_emails()
    print(f"   Found {len(threads)} email threads\n")
    
    # 2. Parse
    records = []
    for i, thread in enumerate(threads):
        print(f"   Parsing thread {i+1}/{len(threads)}: {thread.get('subject', '')[:60]}")
        try:
            records.append(parse_email_thread(thread))
        except Exception as e:
            print(f"   ⚠️  Skipped thread {thread.get('thread_id')}: {e}")
    
    # 3. Deduplicate
    records = deduplicate(records)
    print(f"\n   → {len(records)} unique applications after deduplication\n")
    
    # 4. Write xlsx
    write_xlsx(records)
    
    # 5. Calendar sync
    records = sync_calendar_events(records)
    
    # 6. Re-write with calendar IDs populated
    write_xlsx(records)
    
    print("\n✅ Pipeline complete.")

if __name__ == "__main__":
    run()
```

---

## Running the Tracker

```bash
# Full run (fetch + parse + write + calendar sync)
python tracker.py

# Or via Claude Code
claude "run the internship tracker pipeline"
```

---

## Incremental Updates

To refresh the tracker without reprocessing everything, add a `--update` flag:

```bash
# Only fetch emails newer than latest date in existing xlsx
python tracker.py --update
```

The `--update` logic should:
1. Load existing `applications.xlsx`
2. Find the most recent `last_updated` date
3. Set `DATE_CUTOFF` to that date
4. Only append/overwrite rows where `email_thread_id` already exists or is new

---

## Manual Override

Add a second sheet `Manual Overrides` to `applications.xlsx` for roles applied to via portals where no confirmation email was sent:

| company | role | role_type | date_applied | status | notes |
|---|---|---|---|---|---|
| Goldman Sachs | Spring Week — IBD | Spring Week | 2024-10-01 | Applied | Applied via portal |

The pipeline will merge these rows, giving them priority over email-parsed data.

---

## Status Colour Legend

| Colour | Status |
|---|---|
| 🟦 Light blue | Applied |
| 🟨 Yellow | Online Assessment |
| 🟩 Light green | First Round Interview |
| 🟩 Mid green | Second Round Interview |
| 🟩 Dark green | Assessment Centre |
| 🟩 **Bold green** | **Offer** |
| 🟥 Red | Rejected |
| ⬜ Grey | Withdrawn |
| ⬜ White | Pending |

---

## What I Need From You

Before running the pipeline, confirm or provide:

1. **Date range** — How far back should the search go? (e.g., September 2024 for current cycle)
2. **Extra keywords** — Any company-specific domains or portal senders to add (e.g., `from:@hirevue.com`, `from:@pymetrics.com`)
3. **Calendar** — Which calendar to add events to: `primary` or a specific calendar like `ryan@…`?
4. **Spreadsheet location** — Where to save `applications.xlsx` locally? (default: current directory)
5. **Status definitions** — Are the 9 statuses above right for you, or do you want to add/rename any?
6. **Roles of interest** — List of companies/programmes you've definitely applied to (optional — helps validate parsing accuracy)
