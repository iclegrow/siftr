# Getting Started with Siftr

Siftr is a Copilot CLI skill that triages your Outlook inbox — it classifies
every email by intent (Action vs Inform) and priority (Urgent / Normal / Low),
applies Outlook categories, and moves low-priority mail out of the way. After
setup, just type `siftr` and you're done.

---

## Prerequisites

Before you begin, make sure you have these installed and working:

| Requirement | Why | How to check |
|---|---|---|
| **Windows PC** | Siftr uses Outlook COM automation (Windows only) | You're on Windows ✓ |
| **Outlook desktop** | Must be installed, signed in, and running | Open Outlook and confirm your inbox loads |
| **PowerShell 5.1+** | Ships with Windows — no install needed | Run `$PSVersionTable.PSVersion` in PowerShell |
| **Node.js** (any recent LTS) | Powers the review and digest browser UIs | Run `node --version` — install from [nodejs.org](https://nodejs.org) if missing |
| **GitHub Copilot CLI** | The AI agent that runs Siftr | You should already have this if someone sent you here |
| **WorkIQ access** | Used to resolve your org chart (manager, directs, peers) | Copilot will prompt you if it can't connect |

---

## Setup (5 minutes)

### 1. Clone the repo

```powershell
cd $env:USERPROFILE
git clone https://github.com/iclegrow/siftr.git
```

If you already have it, `git pull` to get the latest.

### 2. Run the install script

```powershell
cd $env:USERPROFILE\siftr
.\install.ps1
```

This checks prerequisites, adds the Siftr module to your PowerShell profile,
and loads it into your current session.

### 3. Launch Copilot and say "siftr"

```
cd $env:USERPROFILE\siftr
copilot
```

Then type:

```
siftr
```

Since this is your first run, Siftr will automatically launch **`siftr setup`**
which walks you through:

1. **Data path** — where to store your personal config (default: `~/.siftr/`)
2. **Org chart** — fetches your manager, direct reports, and peers via WorkIQ
3. **Outlook folders** — creates `Inbox/LowPri` and `Inbox/Meetings` (or
   names you choose)
4. **Categories** — sets up `Urgent`, `Action`, and `Inform` categories (or
   names you choose)
5. **Preferences** — low-priority behavior (move to folder, categorize only,
   or do nothing)

Everything is saved to `config.json` in your personal data directory. You can
re-run `siftr setup` anytime to change settings.

---

## Your First Triage

After setup completes, Siftr immediately triages your inbox. You'll see a
terminal briefing like this:

```
📬 Siftr — 42 emails triaged (since last 24 hours)

🔴 URGENT ACTION (2)
  1. [Jane Smith] "Budget approval needed" — automated approval, deadline today
  2. [Bob Lee] "Prod incident P1" — you're on To, direct ask

🟠 ACTION NEEDED (8)
  1. [Your Manager] "Follow up on Q3 plan" — manager asking for input
  ...

🟢 INFORMED (18)
  ⬆ [HR Partner] "Benefits enrollment reminder" — priority informed
  ...

⚪ LOW PRIORITY (14)
  ...

🏷️📦 Siftr actions: 2 → Urgent, 8 → Action, 18 → Inform, 14 → LowPri
```

Outlook categories are applied automatically. Low-priority mail moves to
`Inbox/LowPri`. Calendar items move to `Inbox/Meetings`.

---

## Review & Learn

After every triage, Siftr opens a **review UI** in your browser at
`http://localhost:8473`. This is where you teach Siftr your preferences:

1. **Scan the list** — each email shows its classification and confidence
2. **Override** any you disagree with (dropdown next to each email)
3. **Add notes** explaining why (optional but helps Siftr learn faster)
4. Click **Save**
5. Back in the CLI, say **`siftr learn`**

Siftr reads your corrections and builds a `rules.md` file with your personal
classification rules. Each triage run gets smarter as your rules grow.

> 💡 **Tip:** The first 2–3 runs will have the most corrections. After that,
> Siftr usually nails it.

---

## Daily Workflow

Once set up, your daily routine is:

| When | What to say | What happens |
|---|---|---|
| Morning | `siftr` | Triages new mail since last run |
| All day | `siftr loop` | Auto-triages hourly until 8pm — set and forget |
| Quick scan | `siftr digest` | Opens a tile-based inbox view in browser |
| After digest | `siftr process my digest` | Marks reviewed emails as read |
| Occasional | `siftr learn` | Processes your review corrections into rules |

---

## All Commands

| Command | What it does |
|---|---|
| `siftr` | Triage inbox (since last scan) |
| `siftr since Monday` | Triage with a custom time window |
| `siftr setup` | Re-run the configuration wizard |
| `siftr refresh org` | Re-fetch your org chart |
| `siftr learn` | Process review feedback into personal rules |
| `siftr digest` | Browser-based inbox digest (unread today) |
| `siftr digest all mails` | Digest including already-read emails |
| `siftr process my digest` | Apply mark-read selections from digest |
| `siftr status` | Show config, org cache age, learning count |
| `siftr dry-run` | Classify without applying any Outlook changes |
| `siftr loop` | Auto-triage hourly until 8pm |
| `siftr stop` | Stop the loop |

---

## Your Personal Files

Everything personal lives in your data directory (default `~/.siftr/`),
**not** in the repo. Nothing to worry about with git.

| File | What it is | You edit it via |
|---|---|---|
| `config.json` | Settings: folders, categories, ports | `siftr setup` |
| `rules.md` | Your classification rules | `siftr learn` |
| `org-cache.json` | Your manager, directs, peers | `siftr refresh org` |
| `last-scan.json` | Timestamp of last triage | Automatic |
| `learnings/*.json` | Review history | Review UI |
| `digests/*.json` | Digest history | Digest UI |

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Outlook COM" error | Make sure Outlook desktop is open and signed in |
| Module not found | Check your PowerShell profile loads `Siftr-Inbox.ps1` |
| Review UI won't open | Check Node.js is installed: `node --version` |
| Org chart empty | Run `siftr refresh org` — needs WorkIQ access |
| Categories not showing | Open Outlook → View → Categories → verify they exist |
| Want to start over | Delete your `~/.siftr/` folder and run `siftr setup` |

---

## Questions?

Ask the person who sent you this, or just type `siftr status` to see if
everything is configured. Siftr is fully self-contained — once set up, it
just works. 🎯
