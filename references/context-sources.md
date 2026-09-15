# Context Sources

Use this reference when the agenda should reflect prior notes, meetings, email, calendar, Slack, Monday.com, or GTD workbook context. Keep searches scoped to the named manager, team, project, or 1:1 folder.

## Recurring meetings: prior-instance & recap sourcing

Every recurring meeting (1:1s, standing syncs) opens with a **Last meeting recap** built from the *previous instance* of the same meeting. Resolving that instance and chaining to it is the core of good follow-up, and it is **not optional** — the renderer rejects an agenda without it. Follow this procedure in order.

### Why the obvious approach does not work

Do **not** try to find the prior note by searching Granola for the counterpart's name. Two independent failure modes make title search unreliable, and both were caught producing silently recap-less agendas:

- **`search_notes` truncates before it filters.** It has no server-side search: it fetches the newest `limit` notes (default **25**) and only then filters on title. Anything older than that window is invisible no matter what it is called. Verified: `search_notes("Vanessa")` returned **0** results at the default limit, while the note existed the whole time and `limit=400` found it.
- **`list_notes` cannot reach older notes at all.** It hard-caps at **200** notes newest-first, and `created_after` does not shift the window — so a monthly meeting's prior instance is routinely outside anything `list_notes` can return. See the Granola section below for the verified numbers and the `granola-index.json` cache that works around it.
- **Granola auto-titles do not match Outlook subjects.** Verified pairs: `1 on 1 with Mike/Dan` → "Mike 1:1"; `Status Update` → "Mariyo 1:1"; `Status Meeting - Dan` → "Initiatives update, board report, and creative delays with Scott"; `Testing and Optimization Review` → "A/B testing review — travel widget, membership offers, and finance homepage with JP".

Correlate on **start time** instead. Titles are corroboration only, never the gate.

### The procedure

1. **Resolve identity, including generic subjects.** Read `Daily Plan\meeting-aliases.json` first. Calendar subjects like `Status Update` or `Status Meeting - Dan` are generic placeholders for a specific person's 1:1; the alias map gives the real counterpart, the agenda title to use, and any tracker path. Key on *normalized subject* (lowercased, trimmed — note real entries have trailing spaces) + `|` + `YYYY-MM-DDTHH:MM`, the same instance-key convention close-day uses.
2. **Resolve the prior instance from Outlook, not from filenames.** Enumerate the series with late-bound COM the way `Get-OutlookMeetings.ps1` does — `GetNamespace('MAPI')` → `GetDefaultFolder(9)`, `Items.IncludeRecurrences = $true`, `Sort('[Start]')`, then `Restrict("[Start] >= '<MM/dd/yyyy hh:mm tt>' AND [Start] <= '...'")` over a 90-day lookback — and take the most recent instance **strictly before** the target meeting. This is authoritative and survives title drift; filename matching does not (real drift on disk: `1300 1 on 1 with Mike-Dan` vs `1300 1-1 with Mike`, `1000 A-B Testing Planning Review` vs `1000 A-B Testing Planning-Review`).
3. **Correlate the Granola note by start time.** Check `Daily Plan\granola-index.json` first — if that instance key is already in `correlations{}`, the note id is settled and no search is needed. Otherwise: a note's `created_at` is the meeting **start in UTC**. Convert the prior instance's local start to UTC and match a note whose `created_at` falls within roughly **−10 to +35 minutes** of the slot — recordings start when the meeting starts, sometimes a few minutes late, essentially never early. When more than one note falls in the window, prefer the nearest, use any title hint to break a tie, and record the match as ambiguous rather than guessing silently. **Derive the UTC offset from the meeting date** (PDT is −7, PST is −8); never hardcode it.
4. **Fall back to the transcript when the summary is null.** Many notes — the direct-report 1:1s especially — have `summary: null`. A null summary is **not** "no note": call `get_transcript` and build the recap from that. Treating a null summary as a miss is one of the ways the recap used to vanish.
5. **Treat the prior agenda `.docx` as a supplementary source, subordinate to the note.** A prior *agenda* is not evidence of a prior *meeting held* — agendas get written for meetings that then move or get cancelled, and an agenda that itself lacks a recap will propagate its own emptiness forward. Verified case: `Agendas\2026_09_14\1300 Dan x Vanessa Monthly Sync.docx` existed for a meeting the calendar placed on 2026-09-15 at 12:00, with no recap of its own. Use the agenda to lift commitments and open loops *after* the note, and only when the calendar confirms that instance actually happened.
6. **Source priority:** Granola note (summary, else transcript) → prior agenda `.docx` → manual / local 1:1 notes → recap email.
7. **Compose the recap** with: a short summary of what was discussed, open follow-ups / action items (with owner where known), decisions made, and suggested talking points derived from last call's loose ends. Unresolved follow-ups must surface as this meeting's talking points or commitments so nothing is dropped.
8. **Record provenance in `recap_sources[]`.** Every agenda payload carries a `recap_sources` array; each entry needs `kind` (`outlook_series`, `granola`, `prior_agenda`, `local_notes`, `tracker`, `email`, `alias_map`), `status` (`found` / `not_found`), and a `detail` or `searched` string. The renderer prints these in the document, so a thin recap reads as *checked and came up empty* rather than unattempted. Also append the resolved match to `correlations{}` in `granola-index.json` (and the note itself to `archived_notes[]` if it came from a deep `search_notes`), so the next run chains for free.
9. **If genuinely nothing is found**, the recap section must contain the exact phrase **"No prior meeting found"** plus the slots and sources searched, and every `recap_sources` entry must carry a `searched`/`detail` string. The renderer enforces this and exits **2** otherwise.

## Local 1:1 Notes

- Prefer a user-provided 1:1 folder when available.
- Search filenames first with Glob, then search content with Grep.
- Prioritize recent agendas, notes, action-item lists, and documents containing the manager's name, `1:1`, `one-on-one`, `commitment`, `blocker`, `feedback`, `priority`, or `growth`.
- For `.docx` notes, extract text with a document-aware tool or unzip/read `word/document.xml` only when no better parser is available.
- Capture only agenda-relevant facts: commitments, outcomes, unresolved questions, explicit feedback, decisions, blockers, asks, and follow-up dates.

## Granola

- Granola is a **required** source for any recurring meeting's recap — not something to reach for only when the user mentions notes. Follow the prior-instance procedure above.
- **Correlate by `created_at` (meeting start, UTC), not by title.** `search_notes` filters on title *after* truncating to the newest `limit` notes, so it silently misses older meetings; see the failure modes above before using it.
- **`list_notes` hard-caps at 200 notes, and `created_after` does not move that window.** Verified 2026-09-15: `list_notes(created_after=2026-07-01, limit=1000)` and `limit=400` both returned exactly `count: 200`, newest-first, reaching back only to 2026-07-23. Anything older than the 200th-newest note is **unreachable through `list_notes` at all** — no limit, cursor, or `created_after` value gets to it. So one `list_notes(limit=200)` call per run covers the whole live window; do not re-query per meeting, and do not expect a bigger `limit` to reach further back.
- **To reach past the 200-note window, use `search_notes` with a title hint and a large `limit`.** It scans deeper than `list_notes` does, but filters on **title only** and truncates before filtering, so the hint has to be right and the limit has to be generous. Verified on the same day: `search_notes("Vanessa", limit=25)` → **0** results; `search_notes("Vanessa", created_after=2026-07-01, limit=400)` → the 2026-07-17 note. Take the hint from `granola_title_hints` in `meeting-aliases.json`.
- **Persist what you find.** `Daily Plan\granola-index.json` holds `archived_notes[]` (notes that have aged out of the live window) and `correlations{}` (resolved instance-key → note id matches). Read it before searching and append to it after a successful deep search, so a monthly meeting can still chain months later instead of re-guessing the title every time.
- **A null `summary` is not "no note."** Call `get_transcript` and build the recap from that. Direct-report 1:1 notes are routinely summary-less.
- Preserve Granola note ids in the detailed notes and in `recap_sources` so a later run can chain deterministically.
- Beyond the recap, useful when the user references transcripts or decisions discussed verbally. Ad-hoc recordings with no calendar block exist (e.g. a "Quick Sync"); include them when they supersede the last scheduled instance, and say so.

## Outlook

- Use calendar search to identify the specific 1:1 date, title, attendee, and recurrence.
- Use email search for prior agendas, follow-ups, commitments, decisions, and manager feedback.
- Read only messages that are plausibly relevant to the agenda.
- Do not send, forward, archive, delete, categorize, or modify Outlook items unless the user explicitly asks.

## Slack

- Use the `slack` MCP for recent manager DMs, project-channel commitments, blockers, decisions, and stakeholder friction.
- Public-channel search is acceptable when relevant. For private channels, DMs, or group DMs, proceed only when the user explicitly asked to include Slack/DM context or gives consent.
- Search narrowly: manager name or user ID, project names, `blocker`, `decision`, `feedback`, `priority`, `ask`, `FYI`, `can you`, `follow up`, and dates around the last 1:1.
- Do not post or draft Slack messages unless the user explicitly asks.

## Atlassian (Jira / Confluence)

- Use the `my-atlassian` MCP when the user asks to include Jira issue, sprint, or Confluence context.
- Search narrowly for the manager name, team name, active project keys, blocked or stale issues, and items assigned to Dan.
- Capture agenda-relevant facts only: owner, status, due date, blocker, decision needed, next milestone, and at-risk work.

## Monday.com

- Use Monday.com when the user asks to include work-board, project, task, status, or blocker context.
- Start from `https://ace-aaa.monday.com/` when a Monday connector or MCP is available.
- The official monday MCP requires the monday marketplace MCP app and OAuth/account permissions; a local package alone is not enough for authenticated board access.
- Search narrowly for the manager name, team name, board names mentioned in local notes, active project names, overdue items, stuck statuses, blockers, and items assigned to Dan.
- Capture agenda-relevant facts only: owner, status, due date, blocker, decision needed, next milestone, and stale or at-risk work.
- If no Monday MCP or authenticated connector is available, note the gap in the agenda context and continue with the other sources.

## GTD Workbook

- Use the GTD workbook as a read-only agenda source unless the user explicitly asks to add or update GTD items (use the `add-gtd-items` skill for writes).
- Default path: `C:\Users\E724101\OneDrive - Automobile Club of Southern California\Dan Yerelian - GTD - $add-gtd-items.xlsx` (the live workbook). The similarly-named `Dan Yerelian - GTD - source for AAA customer acquisition funnel.xlsx` in the home directory is a stale seed copy — do not read it.
- Prefer Excel-native readers or `openpyxl` for inspection. Do not write raw OpenXML parts.
- Review Projects, Waiting For, Next Actions, and Inbox for items involving the manager, the team, the meeting date, high priority work, blockers, open waiting items, and stale next actions.
- Convert workbook rows into agenda prompts: what to clarify, what to unblock, what to deprioritize, what needs a commitment, and what can be closed.

## Synthesis Rules

- Prefer the newest source when priorities conflict.
- Treat private notes and raw message content as source material for the detailed agenda, not as text to send ahead.
- Separate "known from sources" from "suggested prompt" when drafting sensitive sections.
- Include a short "Context reviewed" section in the Word document when context was gathered.
