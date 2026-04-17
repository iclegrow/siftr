# Copilot Instructions — Siftr

This repository contains Siftr, an email triage skill for Copilot CLI.

## Skill location

The skill definition is at `.github/skills/siftr/SKILL.md`. It activates
when the user says "siftr", "triage my email", etc.

## Repository structure

- `.github/skills/siftr/SKILL.md` — Universal skill logic and rules
- `modules/Siftr-Inbox.ps1` — PowerShell COM backend (Outlook automation)
- `review-server/` — Interactive learning review UI (Node.js, port 8473)
- `digest-server/` — Inbox digest UI (Node.js, port 8474)
- `docs/getting-started.md` — New user guide

## Personal data

Personal data (config, rules, org cache, learnings) lives outside the repo
in a user-specific directory discovered at runtime:
1. `$env:SIFTR_PERSONAL`
2. `~/.siftr/`
3. Legacy: `OneDrive/AI-Tools/siftr_personal/`

## When working on Siftr

- SKILL.md uses `$SiftrRoot` for all paths — never hardcode absolute paths
- Siftr-Inbox.ps1 must be saved with UTF-8 BOM for PowerShell 5.1 compatibility
- Config and rules are personal — keep universal rules in SKILL.md, personal
  rules in the user's `rules.md`
- Test changes with `siftr dry-run` before applying to real Inbox
