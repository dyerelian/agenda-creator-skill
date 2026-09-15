---
name: agenda-creator
description: Create agendas and prep documents for any meeting — manager 1:1s, team meetings, project syncs, stakeholder and cross-functional meetings. Use when the user asks to create an agenda, prepare for a meeting, 1:1, or one-on-one, turn prior notes into talking points, gather context from prior agendas/meetings, Granola, Outlook, Slack, Atlassian, Monday.com, or the GTD workbook, or produce a Word document with detailed notes plus concise send-ahead bullets.
---

# Agenda Creator

## Overview

Create a Word agenda for any meeting — a manager 1:1, team meeting, project sync, or stakeholder/cross-functional meeting — that helps the user align expectations, surface truth early, and compound trust over time. Manager 1:1s are the most common case and have a default structure below; for other meeting types, adapt the structure to fit the meeting's purpose and attendees. Regardless of type, keep one short section safe to send to attendees, and make the rest of the document a robust private prep brief with detailed notes, evidence, prompts, and proposed asks.

## Workflow

1. Identify the counterpart, meeting date, meeting time, and meeting type. Ask a concise question only when the counterpart or meeting identity cannot be inferred. Determine the output path using the rules in "Output location & filename" below.
2. Gather relevant context before drafting — for **every** meeting type, not just 1:1s. Review prior agendas or notes for this meeting or counterpart, the most recent meeting(s) on this topic, and any related email threads, notes, or documents. Read `references/context-sources.md` and use Granola, Outlook, Slack, Atlassian, Monday.com, and GTD workbook context within the user's requested scope. **Granola and the Outlook series are not optional** for a recurring meeting's recap — see step 4. For a general (non-1:1) meeting, also look for the meeting's purpose, attendees, and any open decisions or action items from the last time the group met.
3. Extract open loops: prior commitments, unresolved asks, blockers, decisions, feedback themes, priority changes, career/growth topics, and relationship or stakeholder risks.
4. Build the **last meeting recap** — mandatory, and enforced by the renderer. Decide recurrence from **Outlook** (`IsRecurring` on the appointment, or the existence of an earlier instance in the series), never from whether a prior agenda filename happens to match: generically-titled slots like `Status Update` and `DXP/CDP Office Hours` are recurring meetings whose filenames drift, and treating a filename miss as "not recurring" is what silently dropped recaps before. Then follow the "Recurring meetings: prior-instance & recap sourcing" procedure in `references/context-sources.md`: resolve the prior instance from the calendar, correlate its Granola note **by start time** (`created_at` is the meeting start in UTC — title search is unreliable), fall back to `get_transcript` when the summary is null, and use the prior agenda `.docx` only as a supplementary source. Produce a recap covering: a short summary of what was discussed, open follow-ups / action items, decisions made, and suggested talking points. Carry unresolved follow-ups forward. Record what you checked in `recap_sources[]`. If nothing is found, the section must say **"No prior meeting found"** and name the slots and sources searched.
5. Draft two layers:
   - Send-ahead bullets: 3-10 bullets, each 5-10 words, written as concise agenda guideposts safe to paste into an invite, email, or Slack message.
   - Detailed notes: robust private preparation material with context, evidence, prompts, decisions, tradeoffs, risks, and proposed asks.
6. Generate the `.docx` with `scripts/create_agenda_docx.py` (see "Generate the Word document" below). Include the send-ahead bullets in the final response so the user can paste them into a calendar invite, Slack message, or email.

## Output location & filename

Save every meeting-specific agenda under the shared Agendas base folder, inside a subfolder named for the **meeting date**:

```
C:\Users\E724101\OneDrive - Automobile Club of Southern California\Daily Plan\Agendas\<YYYY_MM_DD>\<HHMM> <Title>.docx
```

Rules:

- **Dated subfolder.** Use the meeting date formatted `YYYY_MM_DD` (e.g., today `2026-07-06` → `2026_07_06`). Create the folder if it does not exist — do not error if it already exists:
  ```powershell
  New-Item -ItemType Directory -Force -Path 'C:\Users\E724101\OneDrive - Automobile Club of Southern California\Daily Plan\Agendas\<YYYY_MM_DD>' | Out-Null
  ```
- **Filename = time, then title.** Time is 24-hour `HHMM` with no separator (9:30 AM → `0930`, 2:00 PM → `1400`). Follow it with a single space, then the meeting title. Example: `0930 1-1 with Mike.docx`.
- **No known time.** If the meeting has no specific/known time, use `0000` as the time prefix (e.g., `0000 1-1 with Mike.docx`).
- **Sanitize for Windows.** Filenames cannot contain `\ / : * ? " < > |`. Replace any colon in the title with a hyphen (e.g., "1:1 with Mike" → `1-1 with Mike`); replace other illegal characters with a hyphen or space as reads best.
- Only override this location if the user explicitly asks for a different path.

## Generate the Word document (verified path)

This is the part to get exactly right on this machine. Write an intermediate JSON file, then run the script with the **explicit Python interpreter path**.

```powershell
& 'C:\Program Files\Python312\python.exe' 'C:\Users\E724101\.claude\skills\agenda-creator\scripts\create_agenda_docx.py' --input '<agenda.json>' --output '<agenda.docx>'
```

Sanity-check the script and see the expected output shape with `--example` (writes a sample doc, no input JSON needed):

```powershell
& 'C:\Program Files\Python312\python.exe' 'C:\Users\E724101\.claude\skills\agenda-creator\scripts\create_agenda_docx.py' --example --output "$env:TEMP\example-agenda.docx"
```

Both print `Wrote <path>` and exit 0. The script uses only the Python standard library — no `pip install` needed.

**Write the intermediate JSON without a BOM.** Use the Write tool (it writes UTF-8 with no BOM). If you must write it from PowerShell, do not use `Set-Content -Encoding utf8` (it adds a BOM); use:

```powershell
[System.IO.File]::WriteAllText('<agenda.json>', $json, (New-Object System.Text.UTF8Encoding($false)))
```

(The script reads JSON as `utf-8-sig`, so a stray BOM is now tolerated, but BOM-less is the clean default.)

### Verify the generated doc

The output is a Word `.docx` (a zip of OpenXML). Confirm it opened and contains the expected text by extracting `word/document.xml`:

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead('<agenda.docx>')
$e = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
$sr = New-Object System.IO.StreamReader($e.Open()); $xml = $sr.ReadToEnd(); $sr.Close(); $zip.Dispose()
(($xml -replace '<[^>]+>',' ') -replace '\s+',' ').Trim()
```

### JSON shape

```json
{
  "title": "1:1 with Manager - 2026-06-11",
  "subtitle": "Prepared agenda",
  "send_ahead_bullets": [
    "Align on top priorities for this week",
    "Review prior commitments and next steps",
    "Discuss feedback and growth opportunities today"
  ],
  "context_reviewed": [
    "Prior 1:1 notes from folder",
    "Granola meeting notes"
  ],
  "sections": [
    {
      "heading": "1. Check-in",
      "items": [
        {
          "label": "Energy/morale",
          "body": "Note the user's current operating context and any capacity constraints."
        }
      ]
    }
  ]
}
```

Each section's `items` may be objects (`label` + `body`), bare strings, or a `{label: body}` map. A section may also carry a `notes` array (rendered as bullets). Top-level `notes` becomes an "Additional notes" heading.

## Agenda Structure

### Always include: Last meeting recap

Every agenda — 1:1 or general — includes a **Last meeting recap** section near the top (right after the check-in / purpose section). **The renderer enforces this and exits 2 without it**, so it cannot be quietly skipped. Build it from the previous instance per the "Recurring meetings: prior-instance & recap sourcing" procedure in `references/context-sources.md` (resolve the instance from the Outlook series, correlate the Granola note by start time, transcript when the summary is null, prior agenda as a supplementary source only). It has these parts:

- **Date + source** — the date of the last instance and which source(s) the recap was built from (Granola note, prior agenda, notes, or email).
- **Recap summary** — a short narrative of what was discussed last time.
- **Open follow-ups / action items** — outstanding commitments carried forward, with owner where known.
- **Decisions made** — key decisions reached last time.
- **Suggested talking points** — agenda items derived from last call's loose ends; these can seed this meeting's send-ahead bullets and discussion topics.

Unresolved follow-ups should surface as this meeting's talking points or commitments so nothing is dropped. If there was no prior meeting, state the exact phrase **"No prior meeting found"** plus the slots and sources searched, so the reader knows it was checked, not skipped.

Record provenance alongside it in a top-level `recap_sources` array. Each entry takes `kind` (`outlook_series`, `granola`, `prior_agenda`, `local_notes`, `tracker`, `email`, `alias_map`), `status` (`found` / `not_found`), and a `detail` (or `searched`) string. The renderer prints these under a **Recap sources** label inside the recap section, which is what makes a genuinely empty recap legible as checked:

```json
"recap_sources": [
  {"kind": "outlook_series", "status": "found",
   "detail": "Prior held instance 2026-07-17 09:00 PT, resolved from the recurring series"},
  {"kind": "granola", "status": "found",
   "detail": "Note not_GLRoxszMsAVb4B at 2026-07-17T16:03:07Z vs slot 16:00Z (+3 min); summary null, transcript used"},
  {"kind": "prior_agenda", "status": "not_found",
   "searched": "Agendas\\*\\* for titles matching 'Vanessa' before 2026-09-15"}
]
```

Set `"doc_type": "agenda"` explicitly on agenda payloads. Payloads with no `send_ahead_bullets` (the 1:1 trackers) are exempt from the check automatically; `doc_type` overrides the guess either way, and `--no-require-recap` forces the exemption for one-off prep docs.

### Manager 1:1 (default structure)

For a manager 1:1, use this structure unless the user asks for a different format:

1. Check-in
   - Energy/morale
   - Anything notable personally/professionally
2. Last meeting recap
   - Date + source
   - Recap summary
   - Open follow-ups / action items
   - Decisions made
   - Suggested talking points
3. Five words
   - My five
   - Manager's five
4. Top priorities
   - Priority 1
   - Priority 2
   - Priority 3
   - What should be deprioritized
5. Commitments from last time
   - Commitment
   - Status
   - Result
   - Next step
6. Decisions / blockers / asks
   - Decision needed
   - Blocker
   - Help requested
7. Feedback
   - What should I keep doing?
   - What should I change?
   - Any performance concerns I should know about?
8. Growth / career
   - Skill I'm developing
   - Evidence needed for next level
   - Opportunity to pursue
9. Team / stakeholder dynamics
   - Relationships to improve
   - Team friction or risks
   - Ways I can help the broader team
10. Commitments before next 1:1
   - I will
   - Manager will
   - Date

### General meeting (team, project sync, stakeholder/cross-functional)

For any non-1:1 meeting, adapt the structure to the meeting's purpose and attendees. A solid default:

1. Purpose & desired outcome
   - Why we're meeting and what "done" looks like for this session
2. Last meeting recap
   - Date + source
   - Recap summary
   - Open follow-ups / action items
   - Decisions made
   - Suggested talking points
3. Attendees & roles
   - Who's in the room and what each owns
4. Context / where things stand
   - Recap from the last meeting on this topic; relevant updates
5. Discussion topics
   - Topic, background, and the specific question or decision for each
6. Decisions needed
   - Decision, options/tradeoffs, and who decides
7. Risks / blockers / dependencies
8. Action items & owners
   - Action, owner, due date
9. Next steps / follow-up meeting

## Quality Bar

- Make the agenda specific to this manager, this meeting, and the latest available context.
- Always include the **Last meeting recap** section — the renderer will reject the agenda otherwise. Resolve the previous instance from the Outlook series, correlate its Granola note by start time, and cover date + source, recap summary, open follow-ups, decisions, and suggested talking points. Surface unresolved follow-ups as this meeting's talking points or commitments so they don't get dropped. If no prior meeting exists, say so explicitly with what you searched, rather than dropping the section.
- Never build a recap from a neighbouring agenda file alone. Confirm against the calendar that the instance actually happened, and cite the Granola note id when there is one.
- Convert vague topics into useful prompts: decision needed, tradeoff, risk, evidence, ask, or commitment.
- Treat only the send-ahead bullets as recipient-ready. The rest of the Word document can be much more detailed and candid.
- Keep sensitive source details out of the send-ahead bullets unless the user explicitly wants them shared.
- Preserve citations or source links returned by meeting tools in the detailed notes when available.
- Mark uncertain claims as "verify" instead of presenting them as fact.
- End with clear owners, dates, and next steps.

## Gotchas (Windows / this machine)

- **Send-ahead bullets are validated.** Each must be 5-10 words and there can be at most 10. Outside that range the script exits 1 with `error: Send-ahead bullets must be 5-10 words each`. Count words before generating, or the run fails.
- **The recap is validated too, and it exits 2, not 1.** A missing/empty recap section, a missing `recap_sources`, or an all-`not_found` recap without the "No prior meeting found" sentinel all fail. Use `--validate-only` to pre-check a payload without writing a file — worth doing across a batch before rendering, so one unresolvable meeting doesn't surface halfway through.
- **Do not invoke the script with bare `py`.** `py scripts\create_agenda_docx.py ...` follows the file's `#!/usr/bin/env python3` shebang, which resolves `python3` to the Microsoft Store app-execution alias and fails with `Python was not found` (exit 9009). Use the explicit interpreter path `& 'C:\Program Files\Python312\python.exe' ...` (verified), or `py -3.12 <script>` which pins the version and ignores the shebang.
- **JSON BOM.** PowerShell's `Set-Content -Encoding utf8` writes a UTF-8 BOM. Use the Write tool or `[System.IO.File]::WriteAllText(..., UTF8Encoding($false))` to write the intermediate JSON. (The script reads `utf-8-sig` so it now tolerates a BOM, but other tools may not.)

## Troubleshooting

- `Python was not found ... Microsoft Store` / exit 9009 → you called the script via bare `py` and it followed the shebang. Use the full `python.exe` path or `py -3.12`.
- `error: Unexpected UTF-8 BOM (decode using utf-8-sig)` → the input JSON has a BOM. Rewrite it without a BOM (see above). (Should not occur with the ported script, which reads `utf-8-sig`.)
- `error: Send-ahead bullets must be 5-10 words each` → fix the bullet word counts (5-10 words, max 10 bullets).
- `error: recap: ...` / exit 2 → the Last meeting recap is missing, empty, or unsourced. Resolve the prior instance per `references/context-sources.md`; if there genuinely is no prior meeting, add the "No prior meeting found" sentinel plus a `recap_sources` entry naming what you searched. Do **not** work around this by deleting `send_ahead_bullets` or setting `doc_type` to something else.
- Rendering a **tracker** and hitting a recap error → pass `--no-require-recap`, or check that the payload really has no `send_ahead_bullets`.
- `error: Provide --input or --example` → pass `--input <agenda.json>` (or `--example`); `--output` is always required.
