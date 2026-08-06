# Context Sources

Use this reference when the agenda should reflect prior notes, meetings, email, calendar, Slack, Monday.com, or GTD workbook context. Keep searches scoped to the named manager, team, project, or 1:1 folder.

## Recurring meetings: prior-instance & recap sourcing

Most recurring meetings (1:1s, standing syncs) should open with a **Last meeting recap** built from the *previous instance* of the same meeting. Resolving that instance and chaining to it is the core of good follow-up. Follow this procedure:

1. **Detect recurrence / identity.** Treat a meeting as recurring when its title or counterpart matches a prior agenda or a prior Granola note. For a stable identity across runs, match on *normalized subject* + start time (the same instance-key convention close-day uses: normalized subject + `|` + `YYYY-MM-DDTHH:MM`), so both skills recognize "the same meeting" consistently.
2. **Find the prior agenda (deterministic — this is the follow-up chain).** `Glob` the Agendas base folder for earlier dated instances of the same title:
   ```
   C:\Users\E724101\OneDrive - Automobile Club of Southern California\Daily Plan\Agendas\<YYYY_MM_DD>\<HHMM> <Title>.docx
   ```
   Pick the most recent dated folder *before* the current meeting date whose filename title matches. Read its text (unzip `word/document.xml`) to lift last time's commitments, decisions, and open loops — these carry forward into this agenda.
3. **Find the prior instance's Granola note.** Use the `granola` MCP (`search_notes` / `recent_notes`) with the counterpart or team name, then correlate to the prior instance by **start-time overlap + fuzzy title match**. Prefer `get_note` (AI summary); use `get_transcript` only when exact wording matters.
4. **Source priority for the recap:** Granola note first → prior agenda `.docx` → manual / local 1:1 notes → recap email. State which source(s) the recap was built from.
5. **Compose the recap** with: a short summary of what was discussed, open follow-ups / action items (with owner where known), decisions made, and suggested talking points for this meeting derived from last call's loose ends. Unresolved follow-ups should surface as this meeting's talking points or commitments so nothing is dropped.
6. **If no prior instance is found**, state "No prior meeting found" so the reader knows it was checked, not skipped.

## Local 1:1 Notes

- Prefer a user-provided 1:1 folder when available.
- Search filenames first with Glob, then search content with Grep.
- Prioritize recent agendas, notes, action-item lists, and documents containing the manager's name, `1:1`, `one-on-one`, `commitment`, `blocker`, `feedback`, `priority`, or `growth`.
- For `.docx` notes, extract text with a document-aware tool or unzip/read `word/document.xml` only when no better parser is available.
- Capture only agenda-relevant facts: commitments, outcomes, unresolved questions, explicit feedback, decisions, blockers, asks, and follow-up dates.

## Granola

- Use the `granola` MCP when the user references meeting notes, transcripts, prior 1:1s, or decisions discussed verbally.
- Prefer natural-language meeting queries (`search_notes`, `recent_notes`) before listing many meetings.
- Useful query terms: manager name, `1:1`, `one-on-one`, `commitments`, `blockers`, `priorities`, `feedback`, `growth`, `career`, `performance`, and key project names.
- Preserve Granola citation links in detailed notes when the tool returns them.
- Use `get_transcript` only when exact wording matters.

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
- Default path: `C:\Users\E724101\OneDrive - Automobile Club of Southern California\Dan Yerelian - GTD - source for AAA customer acquisition funnel.xlsx`.
- Prefer Excel-native readers or `openpyxl` for inspection. Do not write raw OpenXML parts.
- Review Projects, Waiting For, Next Actions, and Inbox for items involving the manager, the team, the meeting date, high priority work, blockers, open waiting items, and stale next actions.
- Convert workbook rows into agenda prompts: what to clarify, what to unblock, what to deprioritize, what needs a commitment, and what can be closed.

## Synthesis Rules

- Prefer the newest source when priorities conflict.
- Treat private notes and raw message content as source material for the detailed agenda, not as text to send ahead.
- Separate "known from sources" from "suggested prompt" when drafting sensitive sections.
- Include a short "Context reviewed" section in the Word document when context was gathered.
