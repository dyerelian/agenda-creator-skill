<#
.SYNOPSIS
    Install the agenda-creator Claude skill into ~/.claude/skills.
.DESCRIPTION
    Creates a directory junction at ~/.claude/skills/agenda-creator that points to
    this repo root, so edits here are live in Claude Code immediately. Restart Claude
    Code afterwards so the skill is picked up.
#>
[CmdletBinding()]
param(
    [string]$SkillsRoot = (Join-Path $env:USERPROFILE '.claude\skills'),
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$link = Join-Path $SkillsRoot 'agenda-creator'
New-Item -ItemType Directory -Force -Path $SkillsRoot | Out-Null

if (Test-Path $link) {
    if (-not $Force) {
        Write-Warning "Skill path already exists (use -Force to replace): $link"
        return
    }
    Remove-Item $link -Recurse -Force
}

New-Item -ItemType Junction -Path $link -Target $PSScriptRoot | Out-Null
Write-Host "Linked: $link -> $PSScriptRoot"
Write-Host ""
Write-Host "Done. Restart Claude Code (or start a new session) to load the skill."
