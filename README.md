# Siftr — Email Triage Assistant

Siftr is a Copilot CLI **skill** that triages your unread email into prioritized
categories so you can focus on what matters.

**Skill location:** `.github/skills/siftr/SKILL.md`
The skill auto-activates when you say "siftr", "triage my email", etc.

## Quick Start

1. Clone and install: `git clone https://github.com/iclegrow/siftr.git && cd siftr && .\install.ps1`
2. Launch Copilot in the `siftr` directory
3. Type **`siftr`** to begin — if this is your first time, `siftr setup`
   runs automatically to configure your personal data path, org context,
   folder preferences, and category names
4. Review the prioritized briefing in the terminal
5. Use Outlook categories such as **Urgent**, **Action**, and **Inform** to
   filter the triaged mail
6. Open the review UI at `http://localhost:8473` to override classifications
7. Say **`siftr learn`** to build personal rules from your corrections

📖 **New user?** See the [Getting Started Guide](docs/getting-started.md) for
a detailed walkthrough from zero to first triage.

## Commands

| Command | What it does |
|---|---|
| `siftr` | Triage Inbox-root mail (since last scan) |
| `siftr since Monday` | Triage with a custom time window |
| `siftr setup` | Interactive first-run configuration |
| `siftr refresh org` | Re-fetch your org chart (manager, directs, peers) |
| `siftr learn` | Process your feedback from the review UI |
| `siftr digest` | Inbox digest — scan today's unread emails in browser |
| `siftr digest all mails` | Digest including already-read emails |
| `siftr process my digest` | Apply mark-read selections from digest |
| `siftr status` | Show org cache age and learning file count |
| `siftr loop` | Start hourly triage loop until 8pm |
| `siftr stop` | Stop the running loop |

## Classification Model — Intent × Priority

Siftr classifies each email along two dimensions:

- **Intent:** Action (you need to do something) or Inform (FYI only)
- **Priority:** Urgent, Normal, or Low

These combine into five tiers plus one routing tier:

| Tier | Intent | Priority |
|---|---|---|
| 🔴 Urgent Action | Action | Urgent |
| 🟠 Action Needed | Action | Normal |
| 🟢⬆ Priority Informed | Inform | Urgent |
| 🟢 Informed | Inform | Normal |
| ⚪ Low Priority | Inform | Low |
| 📅 Calendar | *(routing)* | — |

Action + Low does not exist; minimum priority for Action is Normal.

## Outlook Integration

- Siftr reads and writes through Outlook COM against the **Inbox root**.
- It can include **read** mail during triage.
- It can skip mail that already has Outlook categories.
- Three Outlook categories are used (names configurable via `config.json`):
  `Urgent`, `Action`, `Inform`
- Urgent items are dual-categorized:
  - **🔴 Urgent Action** → `Action` + `Urgent`
  - **🟢⬆ Priority Informed** → `Inform` + `Urgent`
- Normal items get their intent category only:
  - **🟠 Action Needed** → `Action`
  - **🟢 Informed** → `Inform`
- Low Priority and Calendar get no category (behavior configurable: move to
  subfolder, categorize only, or do nothing)
- Default folder moves:
  - **⚪ Low Priority** → `Inbox/LowPri`
  - **📅 Calendar** → `Inbox/Meetings`

## Personal Data & Configuration

Siftr stores personal data in a configurable location (discovered in order):
1. `$env:SIFTR_PERSONAL` environment variable
2. `~/.siftr/` (default)
3. Legacy: `OneDrive/AI-Tools/siftr_personal/`

### Three personal files

| File | Purpose | Managed by |
|---|---|---|
| `config.json` | Mechanical settings: folders, categories, ports, org domain | `siftr setup` |
| `rules.md` | Personal classification rules (people, DLs, topics) | `siftr learn` |
| `org-cache.json` | Manager, direct reports, peers, and cached SLT | `siftr refresh org` |

### First-run experience

New users run `siftr setup` (auto-triggered on first use) which walks through:
prerequisites check, data path selection, org resolution via WorkIQ, folder
and category configuration, and Outlook folder creation.

After the first triage run, `siftr learn` creates personal rules from your
corrections. Each run gets better as `rules.md` grows.

## Folder Structure

```
siftr/                            ← repo root
  .github/
    skills/siftr/SKILL.md         - Universal skill logic and rules
    copilot-instructions.md       - Copilot repo context
  modules/
    Siftr-Inbox.ps1               - Outlook COM backend (config-aware)
  review-server/                  - Interactive learning review UI
    server.js                     - Zero-dep Node.js HTTP server (:8473)
    public/index.html
  digest-server/                  - Inbox digest UI
    server.js                     - Zero-dep Node.js HTTP server (:8474)
    public/index.html
  scripts/
    Start-SiftrFullLoop.ps1       - Standalone full-loop runner with recovery
  docs/
    getting-started.md            - New user guide
  README.md                       - This file
  CHANGELOG.md                    - Change history
  install.ps1                     - First-time setup script

Personal data directory (e.g. ~/.siftr/):
  config.json                     - User configuration
  rules.md                        - Personal classification rules
  org-cache.json                  - Cached org chart
  last-scan.json                  - Last triage timestamp
  learnings/                      - Triage review JSON files
  digests/                        - Digest JSON files
  loop-state.json                 - Loop mode state + runner heartbeat metadata
```

## Learning Mode

After each triage run, **all** classifications are exported to a JSON file in
`siftr_personal/learnings/`. A local review server launches automatically at
`http://localhost:8473` — open it in your browser to see every email with its
classification, confidence level, and your addressing context (To/CC/DL).

Override any classification you disagree with using the dropdown, add optional
notes, then click **Save**. When you're done reviewing, say `siftr learn` in
the CLI to process your corrections and update the priority rules.

The review server code lives in `review-server/` (zero-dependency
Node.js).

## Requirements

- **Windows PC** with **Outlook desktop** installed and running (COM automation)
- **PowerShell 5.1+** (Windows PowerShell)
- **Node.js** (any recent LTS) — for review and digest servers
- **Copilot CLI** with WorkIQ access — for org chart resolution

## Installation

```powershell
git clone https://github.com/iclegrow/siftr.git
cd siftr
.\install.ps1
```

Then open Copilot CLI and say **`siftr`** — the setup wizard runs automatically
on first use. See [docs/getting-started.md](docs/getting-started.md) for the
full walkthrough.
