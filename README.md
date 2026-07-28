# Agenda Creator Skill

A Claude Code skill that creates Word agendas and prep documents for any meeting —
manager 1:1s, team meetings, project syncs, and stakeholder/cross-functional meetings.

Each agenda has two layers:

- **Send-ahead bullets** — 3–10 concise bullets (5–10 words each) safe to paste into a
  calendar invite, email, or Slack message.
- **Detailed prep brief** — a robust private document with context, evidence, prompts,
  decisions, risks, and proposed asks.

Every agenda also includes a **Last meeting recap** section (date of the last meeting +
top 3 points discussed).

## Layout

- `SKILL.md` — the skill instructions Claude Code loads.
- `scripts/create_agenda_docx.py` — builds the `.docx` from an intermediate JSON file
  (Python standard library only; no `pip install` needed).
- `references/context-sources.md` — how to pull context from Granola, Outlook, Slack,
  Atlassian, Monday.com, and the GTD workbook.

## Install

This repo is consumed as a Claude Code skill via a directory junction at
`~/.claude/skills/agenda-creator` pointing to this repo root:

```powershell
./install.ps1
```

Restart Claude Code (or start a new session) afterwards so the skill is picked up.

## Updating

The live skill at `~/.claude/skills/agenda-creator` is a junction into this repo, so
edits to `SKILL.md` / `scripts` here are live immediately. Commit and push to back them up:

```powershell
git add -A; git commit -m "..."; git push
```
