# Context Sources

Use this reference when the agenda should reflect prior notes, meetings, email, calendar, Slack, Monday.com, or GTD workbook context. Keep searches scoped to the named manager, team, project, or 1:1 folder.

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
