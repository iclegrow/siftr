# Siftr Changelog

All notable changes to the Siftr email triage skill are documented here.
This log is maintained by Copilot to preserve context across sessions.

---

## 2026-04-30 — Loop mode now skips auto-digests and records classifier drift

- **Digest stays manual while loop mode focuses on triage**: `siftr loop` no
  longer shuts down or launches the digest server on its own. Digest capability
  is still available through the manual `siftr digest` flow.
- **Hourly loop contract is simpler and safer**: the scheduler now runs triage,
  review-store updates, and state persistence only, reducing cross-feature
  coupling during long unattended runs.
- **Heuristic-vs-LLM comparison metrics are now recorded**: each successful LLM
  triage decision is shadow-compared against the local heuristic classifier and
  `loop-state.json` now accumulates sample count, exact-match count, tier
  disagreements, and action/urgent boundary disagreements.
- **Loop triage now ignores historical Inbox backlog**: hourly loop cycles and
  one-off recovery cycles are capped to the most recent 72 hours of mail, even
  when older Inbox-root messages remain uncategorized. This keeps pre-Siftr
  backlog out of the loop's critical path and focuses automation on current
  mail only.
- **Thread loading no longer rescans the Inbox once per conversation**: loop
  triage now builds a single per-cycle conversation index for the active
  threads, then enriches those records in place. This removes the worst
  repeated Inbox-scan multiplier from `thread-loading`.
- **One-off recovery can now target an explicit window and batch size**:
  `Start-SiftrFullLoop.ps1 -RunOneCycleNow` accepts an override `Since` and
  message limit, which makes guarded drain runs like "all of today" possible
  without mutating the normal loop bookmark logic first.
- **Review page columns now stay laptop-friendly**: the review table uses a
  fixed layout with tighter column caps and ellipsis on long sender/subject
  text so one malformed title can no longer blow the whole grid off screen.
- **Review titles now recover from corrupted stored text**: the review UI fixes
  obvious mojibake on load, and the review-store merge path now refreshes
  suspicious subject/sender fields from Outlook by `InternetMessageId` instead
  of guessing at a lossy text conversion.

---

## 2026-04-29 — Manual backlog recovery now completes reliably

- **After-hours one-off recovery runs no longer corrupt loop state**:
  `New-LoopState` now suppresses helper log output so `loop-state.json` stays a
  single state object instead of accidentally becoming an array.
- **Loop review path handling is normalized**: review JSON paths are now reduced
  to one canonical path before writing files or launching the review server, so
  duplicated path values cannot crash startup.
- **Manual recovery batches are bounded and practical**: one-off recovery runs
  now process a smaller inbox slice per pass and use the local heuristic
  classifier directly, avoiding long LLM stalls while clearing backlog.
- **Manual recovery no longer hangs in digest generation**: after a one-off
  recovery cycle succeeds, the runner finalizes immediately instead of launching
  digest work for the same after-hours pass.

---

## 2026-04-29 — Zero-result diagnostics and guarded recovery

- **Zero-email work-hour cycles now self-check**: when a loop cycle returns
  zero messages during the 8 AM-8 PM workday window, Siftr now runs a bounded
  diagnostic against a wider "today since 12:01 AM" Inbox-root window before
  trusting the zero result.
- **Scan anomalies are explicit**: if the primary scan returns zero but the
  fallback diagnostic still finds unread or uncategorized Inbox-root mail,
  Siftr emits `scan_anomaly` events, logs a warning, and immediately retries
  triage with the wider bounded window.
- **Loop health state now tracks quiet periods**: `loop-state.json` records
  `consecutiveZeroCycles`, `lastNonZeroCycleAt`, `lastDiagnosticAt`,
  `lastDiagnosticResult`, and `lastFallbackCount` so repeated zero-email work
  cycles can be distinguished from a genuinely quiet inbox.

---

## 2026-04-29 — Inbox backlog fix: uncategorized mail survives restarts

- **Uncategorized Inbox backlog is no longer hidden by `last-scan.json`**:
  `Get-SiftrInboxRootMessages -SkipCategorized` now keeps older uncategorized
  Inbox-root mail eligible even when it predates the bookmark.
- **Restart-safe triage behavior**: after a crash or restart, Siftr now keeps
  seeing previously unprocessed Inbox-root mail until it is actually categorized
  or moved, instead of silently skipping it because the bookmark advanced.

---

## 2026-04-29 — Loop classifier diagnostics: fix false LLM-failure labeling

- **Copilot wrapper hardened for Windows PowerShell 5.1**: loop mode now trusts
  the CLI's structured `result.exitCode` event instead of relying only on the
  polled `Process.ExitCode`, which can be blank in the detached runner even when
  the Copilot call actually succeeded.
- **Learning records now separate classification source from diagnostics**:
  review rows now include `classificationSource`, `diagnosticCode`, and
  `fallbackNote`, so user-facing fields no longer imply an LLM failure when the
  fallback path was triggered by a runtime issue.
- **`uncertainty` is now classification-only**: loop review rows only populate
  `uncertainty` for low-confidence LLM decisions, not for heuristic fallback or
  process-wrapper problems.
- **Archived loop logs now live under `siftr_personal\Logs`**: daily rotation
  and retention now move dated `loop-run.*` / `loop-events.*` files into a
  dedicated `Logs` subfolder, and legacy root-level archives are migrated there
  during normal cleanup writes.

---

## 2026-04-29 — Universal rule: SLT mail is always priority informed

- **New universal Phase 1 rule**: mail from **CEO Person** or one of his
  direct reports now automatically classifies as **PRIORITY INFORMED**.
- **Org cache schema extended**: `org-cache.json` now documents an `slt`
  section with `ceo` and `directs`, so the rule can be driven from cached
  org data instead of hard-coded mail addresses.
- **Loop fallback logic updated**: the local heuristic classifier now treats
  `slt` senders as a high-confidence Phase 1 `PRIORITY INFORMED` decision, and
  prompt records expose `senderRelation = slt` so the LLM path sees the same
  relationship.

---

## 2026-04-29 — Loop learning: always-on daily review feed

- **Loop mode now keeps learning active**: each hourly cycle merges its
  classifications into a day-scoped review JSON in `siftr_personal/learnings/`
  instead of skipping learning export entirely.
- **One row per unique email**: loop review data is keyed by unique message ID,
  so the same message is updated in place if reprocessed while true resends
  still appear as separate rows.
- **Review server is now kept alive during loop mode**: the loop initializes
  the review file at startup, launches `review-server` against it, and
  re-starts the server automatically if it is shut down mid-day.
- **Daily reset is automatic**: the active loop review file rolls over to a new
  date-named JSON at the start of the next day's loop, effectively clearing the
  live review cache without destroying prior daily history files.

---

## 2026-04-27 — Loop hygiene: daily log rotation and 7-day retention

- **Loop logs now rotate automatically by day**: `loop-run.log` and
  `loop-events.jsonl` are rolled to dated archive files on the first write of a
  new day instead of growing forever as single append-only files.
- **Seven-day retention added**: archived loop logs older than 7 days are
  deleted automatically during normal loop writes, keeping the personal data
  folder tidy without affecting the live `loop-state.json` / `last-scan.json`
  files.
- **No scheduler dependency on old logs**: rotation is implemented directly in
  the append helpers, so loop execution still treats the logs as write-only
  diagnostics rather than reading large history files back into the scheduler.

---

## 2026-04-24 — Reliability: degraded mode, failure counters, and loop events

- **Scheduler is no longer fail-stop**: cycle failures and digest failures are
  now caught and recorded without killing the outer loop. The runner preserves
  state, logs the error, and schedules the next cycle instead of exiting.
- **Per-item degraded mode added**: when an individual Copilot classification
  fails, the loop now falls back to a local heuristic classifier for that
  thread or digest item instead of failing the whole cycle.
- **Bad items are quarantined instead of killing the pass**: if both the LLM
  path and fallback path fail for a specific message, the loop records that
  item as quarantined and continues processing the rest of the Inbox.
- **Structured event log added**: the runner now writes `loop-events.jsonl`
  alongside `loop-run.log` so degraded decisions, quarantined items, phase
  failures, and lifecycle transitions are machine-readable.
- **Richer runtime state added**: `loop-state.json` now tracks
  `leaseExpiresAt`, `consecutiveFailures`, `lastSuccessfulCycleAt`,
  `lastFailureAt`, `lastFailurePhase`, `degradedModeCount`, and
  `quarantinedCount` for better recovery and diagnostics.
- **Heartbeat stays alive during long Copilot calls**: the subprocess wait loop
  now refreshes the loop heartbeat while the CLI is running, reducing false
  stale-runner detection during long LLM work.

---

## 2026-04-23 — Reliability: loop ownership, atomic state writes, crash cleanup

- **Loop owner metadata added**: `loop-state.json` now records the active
  runner's `runnerId`, `pid`, `parentPid`, `host`, `scriptPath`, and a rolling
  `heartbeatAt` so Siftr can tell a live loop from an abandoned one.
- **Duplicate full loops prevented**: a new `Start-SiftrFullLoop.ps1` instance
  now detects other live copies of the same script and exits instead of racing
  an already-running loop.
- **Abandoned state resumes immediately**: if `loop-state.json` is still
  `active` but no other loop runner is alive, the next launch now resumes that
  state immediately instead of waiting for the old 90-minute stale window to
  expire.
- **Crash / unexpected-exit cleanup**: the runner now marks the loop
  `stopped` with `stopReason` / `lastError` if it crashes or exits while the
  state is still active, avoiding the exact "active at 9:00 AM with no agent"
  failure mode.
- **Atomic state writes**: loop state JSON is now written via temp-file replace
  instead of direct overwrite, which is safer in a OneDrive-synced personal
  folder.
- **Copilot subprocess timeout**: the full-loop runner now kills and fails a
  stuck `copilot` classification subprocess after a bounded timeout instead of
  hanging forever mid-cycle with `status: "active"`.
- **Digest folder-name crash fixed**: the digest path no longer uses an invalid
  inline `try { } catch { }` expression while reading `item.Parent.Name`, which
  had been crashing the loop immediately after a successful triage cycle when a
  pending digest slot was processed.
- Source: the 2026-04-23 full LLM loop resumed on top of older state, ran an
  8:54 AM cycle, then left `loop-state.json` active with `nextCycleAt` at
  9:00 AM after the background agent was gone.

---

## 2026-04-22 — Loop: default final cycle moved to 8 PM

- **Later default stop time**: `siftr loop` now defaults to **8:00 PM local**
  so the final scheduled triage cycle runs at or just after 8 PM before the
  loop stops for the day.
- **Local runner path fixed**: the detached local loop runner now points at the
  standalone `<local-path> checkout instead of the removed
  `copilot-home\siftr` submodule path.

---

## 2026-04-22 — Security: tighter prompt-boundary handling

- **Prompt-injection guidance added**: SKILL.md now explicitly treats email
  bodies, replies, and summaries as untrusted content and forbids treating
  them as instructions that can modify Siftr policy or runtime behavior.
- **Category overrides constrained**: `Invoke-SiftrInboxActions` now ignores
  arbitrary category overrides by default. A caller must explicitly mark an
  override as trusted, and even then the override must match the category set
  already allowed for the chosen tier.
- **Digest HTML sanitized**: the browser digest UI now sanitizes `summaryFull`
  with a small allow-list of formatting tags instead of rendering arbitrary
  raw HTML from digest JSON.
- **Digest action text clarified**: the UI now labels `actionText` as a
  user-authored field, and SKILL.md documents that it must start empty and
  must never be auto-populated from email content.

---

## 2026-04-21 — Reliability: safer Inbox fetch + safer bookmark advancement

- **Inbox fetch hardened**: `Get-SiftrInboxRootMessages` and
  `_Get-SiftrConversationInboxItems` now snapshot Outlook folder items with
  `GetFirst()` / `GetNext()` before filtering instead of relying on direct
  `foreach` over the live COM collection.
- **No early stop on older items**: the root fetch no longer `break`s on the
  first message older than `Since`; it continues scanning the snapshot so a
  misordered or partially refreshed COM view cannot hide newer uncategorized
  mail later in the folder.
- **Diagnostics improved**: inbox records now include `ReceivedTime` and
  `Categories`, which makes missed-mail audits match Outlook reality.
- **Loop guidance tightened**: SKILL.md now requires capturing a
  fetch-start bookmark before each cycle and writing that value back to
  `last-scan.json` only after successful actions. Loop mode also retries
  suspiciously tiny fetches once before declaring the Inbox drained.
- **Mixed-thread calendar safety**: conversation-wide fan-out now preserves
  `CALENDAR` routing for meeting-class items even when the latest non-meeting
  reply in that thread is classified as Action or Inform.
- Source: a 3 PM loop cycle under-fetched the Inbox and a cycle-end bookmark
  advanced past messages that arrived during the run, leaving uncategorized
  Inbox mail behind.

---

## 2026-04-20 — Learning: soft asks in reply threads → ACTION NEEDED

- **Expanded Phase 2 🟠 rule**: polite / indirect asks now count as real asks
  when they are directed to the user on the **To** line. Examples:
  **"let us know if you are good to..."**, **"would you be open to..."**,
  **"can you help kick off..."**, **"happy to lean in from there"**.
- **Expanded direct-ask rule**: short question-form asks for missing info or
  approval (for example **"Do you have a phone number for...?"** or
  **"Are you OK with...?"**) now explicitly count as 🟠 ACTION NEEDED.
- **Loop-mode clarification**: hourly triage must not rely on subject-only
  shortcuts for reply threads. If the message is a reply/forward or may contain
  a buried ask, read the latest body text and apply full Phase 2 rules.
- Source: Sean Morgan thread (`RE: Feedback Requested: W+D Creator Council -
  Gaming Focus`) was incorrectly classified 🟢 INFORMED even though the body
  clearly asked User to confirm ownership of the Xbox follow-up.
- Source: Marjorie Ferris thread (`RE: Package to send`) was incorrectly
  classified 🟢 INFORMED even though the body asked User for Project Contact's phone
  number.

---

## 2026-04-20 — Fix: meeting requests now participate in CALENDAR routing

- **Inbox fetch expanded**: `Get-SiftrInboxRootMessages` now includes Outlook
  meeting items (`IPM.Schedule.Meeting*`), not just `olMail`.
- **Message metadata expanded**: fetched items now include `MessageClass` so the
  skill can distinguish normal mail from meeting requests / updates /
  cancellations.
- **Calendar routing clarified**: any item with `MessageClass` beginning
  `IPM.Schedule.Meeting` should classify as **📅 CALENDAR** and follow the
  configured `CALENDAR -> Meetings` move rule.
- Source: Inbox meeting items like `Windows + Xbox Weekly LT Scrum` and
  `FY27 budget model` were staying in Inbox because they were never entering the
  Siftr pipeline.

---

## 2026-04-21 — Conversation-wide labeling from latest thread state

- **New helper**: `Get-SiftrConversationRootMessages` returns all Inbox-root
  items for a given `ConversationId`, so Siftr can classify the latest message
  with real sibling-thread context instead of relying only on quoted history.
- **Action fan-out**: `Invoke-SiftrInboxActions` now treats `ConversationId` as
  a conversation-level decision. It picks the latest classified item in the
  batch for that conversation, then applies that tier to all currently
  uncategorized Inbox-root siblings in the same thread.
- **No retroactive rewrites**: already-categorized earlier messages are left
  alone, so later thread turns can escalate or de-escalate without changing
  historical labels.
- Source: uncategorized earlier replies like `Re: 1:1 Michael/User` and
  `RE: CP+ Born Green - 4/20/2026` were being skipped because only the newest
  thread item received actions.

---

## 2026-04-17 — Shareability Refactor: Personal Rules & Config

**Major refactor** to make Siftr usable by anyone, not just the original author.

- **New §0 Load Configuration**: Siftr now discovers personal-data directory
  via `$SIFTR_PERSONAL` → `~/.siftr/` → legacy OneDrive path. Loads
  `config.json`, `rules.md`, and `org-cache.json` from there.
- **Personal rules extracted to `rules.md`**: ~25 personal classification rules
  (named people, DLs, domains, branch patterns, priority topics, trusted
  vendors) moved from SKILL.md to the user's `rules.md` file. SKILL.md now
  contains only ~15 universal rules that work for any Outlook user.
- **New `config.json`**: Mechanical settings (folder names, category names,
  action behaviors, org domain, server ports) are now configurable via JSON
  instead of hardcoded in `Siftr-Inbox.ps1`.
- **New `siftr setup` command (§12)**: Interactive walkthrough for new users —
  checks prerequisites, sets up data path, resolves org context, configures
  folder/category preferences, creates Outlook folders.
- **Siftr-Inbox.ps1 refactored**: Config discovery at module scope, folder
  and category rules loaded from config.json with legacy defaults as fallback.
  New public functions: `Get-SiftrPersonalPath`, `Get-SiftrConfig`.
  `_Get-InboxSubfolder` now supports `-AutoCreate` for setup.
- **§7 (siftr learn) retargeted**: Learning workflow now edits `rules.md`
  instead of SKILL.md. Universal rules never modified through learning.
- **All hardcoded paths replaced** with config-relative references throughout
  SKILL.md §1–§11.
- **Backward compatible**: Legacy OneDrive path still works. No config.json
  needed for existing users (falls back to legacy defaults).

---

## 2026-04-17 — Feature: Loop Mode (§11)

- **New `siftr loop` command** — hourly automated triage running on the hour
  until 7pm (configurable). Full Phase 1 + Phase 2 LLM classification each
  cycle, brief one-liner output per cycle, auto-digest at noon and 5pm.
- Loop state persisted to `siftr_personal/loop-state.json` for resilience
  across session restarts and `/compact` calls.
- Digest slot tracking prevents skips and duplicates.
- Added `siftr stop` command to end loop gracefully.
- Added to §8 commands table: `siftr loop`, `siftr loop until {time}`,
  `siftr stop`.

---

## 2026-04-17 — Learning: Event/conference surveys → ACTION NEEDED

- **New Phase 1 🟠 rule**: Event or conference surveys (e.g., WinHEC onsite
  survey) are now classified as 🟠 ACTION NEEDED — even when sent via a
  distribution list. These have a clear call to action (complete the survey).
- Source: user override on WinHEC Onsite Survey during digest review.

---

## 2026-04-16 — Learning: EOD Coverage Reports + direct-report FYI forwards → INFORMED

- **New Phase 1 rule**: Subject matching **"EOD Coverage Report"**, **"End of
  Week"**, **"Weekly Coverage"**, or similar comms-summary patterns is now 🟢
  INFORMED even when the user and manager are both on the To line. Exception:
  if the email is a reply/forward where the user is directly mentioned by name,
  Phase 2 may escalate.
- **New Phase 2 clarification**: Direct report FYI forwards with
  **importance: low** and no ask → 🟢 INFORMED, not 🟢⬆. The low-importance
  flag is the sender's own signal that no action is expected.
- Both overrides came from low-confidence classifications in the 2026-04-16
  triage run (17 emails). User notes emphasized that these are routine
  broadcasts / informational forwards with no call to action.

---

## 2026-04-14 — Learning: M365 doc-comment notifications → LOW PRIORITY

- **New Phase 1 rule**: Subject matching the M365 document-comment pattern
  (`"<Person> left a comment in <Doc>"` or `"<Person> replied to a comment in
  <Doc>"`) is now classified ⚪ LOW PRIORITY regardless of sender or document.
- Previously these were classified 🟢⬆ PRIORITY INFORMED. User overrode 15
  instances in one batch — these are automated SharePoint notifications, not
  substantive enough for elevated priority.
- Also: W+D Town Hall Recap upgraded from ⚪ → 🟢 as a one-off (not a rule
  change; user noted it as "likely unique").

## 2026-04-14 — Digest: exclude handled approvals

- Added digest exclusion rule: emails from **MSApprovalNotifications** that
  are already **read** are omitted from the digest (assumed handled). Unread
  approval emails are still shown.

## 2026-04-10 — Inbox Digest feature

- **New feature: `siftr digest`** — tile-based browser UI for fast inbox scanning.
- Tiles grouped by tier (Urgent → Low), each section collapsible.
- Each tile shows: subject, sender, addressing (To/CC/DL), tier badge, and a
  2-line AI-generated summary. Click to expand full bulleted summary with
  highlighted keywords and action items.
- **Mark Read toggle** on each tile — auto-saves state. "Mark Unread" on hover
  to undo. Already-read emails appear dimmed with no toggle.
- `siftr process my digest` reads saved selections and marks messages + full
  conversations as read in Outlook via COM.
- Digest server runs on port 8474 (separate from review server on 8473).
- Added `Set-SiftrMessageReadState` to `Siftr-Inbox.ps1` — marks messages and
  optionally entire conversations as read/unread via Outlook COM.
- Digest JSON stored in `siftr_personal/digests/`.
- Added SKILL.md §9 (Digest Mode) and §10 (Process Digest).

### Digest fixes (same day)
- **Accordion auto-sizing:** expanded tiles now use `max-height: none` after
  animation so summaries are never cut off regardless of length.
- **LOW PRIORITY excluded:** digest UI filters out ⚪ LOW PRIORITY emails.
- **Thread count badges:** tiles in multi-message conversations show a purple
  "N in thread" badge (e.g., "10 in thread" for D.E.Shaw).
- **Subfolder scanning:** added `-IncludeSubfolders` to
  `Get-SiftrInboxRootMessages` so emails moved by prior triage (LowPri,
  Meetings folders) are still captured in the digest.
- Added `IsRead` field to inbox message output.
- Fixed syntax error (extra closing brace) in Siftr-Inbox.ps1.

---

## 2026-04-10 — Interactive learning review UI (replaces CSV)

- **Replaced CSV learning export with interactive browser-based review UI.**
- After triage, ALL classifications (not just low-confidence) are exported to
  a JSON file in `siftr_personal/learnings/`.
- A zero-dependency Node.js server (`siftr/review-server/`) launches at
  `http://localhost:8473` serving a dark-themed "Signal Deck" UI.
- The UI shows: subject, sender, how user was addressed (To/CC/DL), tier badge,
  confidence level, override dropdown, and notes field.
- Features: tier filter buttons, search, "modified only" toggle, Ctrl+S save,
  expandable row details (reason, uncertainty, recipients).
- `siftr learn` now reads from JSON files (with legacy CSV fallback), shuts
  down the review server, and processes user overrides and notes.
- Updated SKILL.md sections 6 (JSON export + server launch) and 7 (JSON-based learn).

---

## 2026-04-10 — Fix UTC→local timezone bug in message fetch (`9d3fb15`)

- `Get-SiftrInboxRootMessages` now auto-converts UTC `$Since` to local time.
- Outlook COM `ReceivedTime` is always local; passing a UTC bookmark caused
  the comparison to break early and return 0 results.
- Root cause: `last-scan.json` stores UTC, but the value was compared raw
  against local-time COM properties.

## 2026-04-10 — Converted to Copilot CLI skill

- Moved instructions from `siftr/siftr.instructions.md` →
  `.github/skills/siftr/SKILL.md` with YAML frontmatter (name, description,
  trigger phrases).
- Copilot CLI auto-discovers the skill via `.github/skills/` convention;
  the manual routing in `.github/copilot-instructions.md` is no longer needed.
- `siftr/` directory retained for documentation (README, CHANGELOG, Plan).
- PowerShell module stays centralized in `modules/Siftr-Inbox.ps1`.

## 2026-04-10 — Learning: Teams join requests → LOW PRIORITY (`051d59a`)

- Added Phase 1 rule: automated "request to join private team" notifications
  from Microsoft Teams → ⚪ LOW PRIORITY (these go to multiple owners).
- Real person replies asking for input can still escalate via Phase 2.
- Source: 1 user override from `siftr-2026-04-10-1214.csv`.
- First triage run with new Intent × Priority model: 40 conversations
  classified (1 🔴, 9 🟠, 10 🟢⬆, 12 🟢, 8 ⚪).

## 2026-04-09 — Intent × Priority model refactor (`e88f12b`)

**Breaking change** — replaced the flat 7-tier classification list with a
two-dimensional Intent × Priority model.

- **New model:** Intent (Action / Inform) × Priority (Urgent / Normal / Low)
  yields 5 tiers + 📅 Calendar routing. Action+Low does not exist.
- **Removed 🟡 AWAITING RESPONSE** — merged into 🟠 ACTION NEEDED; "awaiting
  response" is now a Reason annotation, not a separate tier.
- **Removed 🔵 EXTERNAL** as a primary tier — external mail goes through normal
  Phase 1 / Phase 2 classification. Spam still routes to ⚪ LOW PRIORITY.
- **Dual-categorization for urgent items:** 🔴 → `Action` + `Urgent`,
  🟢⬆ → `Inform` + `Urgent`. Normal items get intent only.
- **3 Outlook categories:** `Urgent`, `Action`, `Inform` (removed `Response`
  and `External`).
- **CSV schema:** added `Intent` and `Priority` columns, renamed old
  `Priority` → `Tier`.
- Updated: `siftr.instructions.md`, `Siftr-Inbox.ps1`, `README.md`,
  `Siftr_Plan.md`.

## 2026-04-09 — Two-phase classification structure (`0f5d367`)

Restructured Section 3 into **Phase 1 (Pattern Rules)** and **Phase 2
(Content Analysis)**. Phase 1 is deterministic metadata matching (sender,
subject, DL). Phase 2 reads the body and applies judgment. All existing
rules preserved; this was a reorganization, not a rule change.

## 2026-04-09 — Categorize all external messages (`3f4e2ec`)

- All external (non-org-domain) messages now go through classification.
- External spam → ⚪ LOW PRIORITY (moved to `Inbox/LowPri`).
- Possibly-legitimate external mail received the `External` Outlook category
  *(later removed in the Intent × Priority refactor)*.

## 2026-04-09 — System consistency audit (`83573c6`)

- Fixed 7+ inconsistencies: stale "24h" references, path mismatches for
  org-cache / learnings / last-scan, missing Calendar tier docs, @mention
  rule conflict.
- Moved @mention rule from 🔴 to 🟠 per user's own learning feedback.
- All paths now reference the OneDrive location
  (`AI-Tools/siftr_personal/`).

## 2026-04-09 — Remove 'Response' Outlook category (`6c7fbfa`)

- Removed `Response` from tier→category mapping in all 4 files.
- Both ACTION NEEDED and AWAITING RESPONSE now map to `Action` only.

## 2026-04-09 — Use last-scan.json as default time window (`4564180`)

- Fixed bug: Section 2 still said "Default window: last 24 hours" even
  though `last-scan.json` was being written.
- Now reads `last-scan.json` on startup; falls back to 24h only if the
  file is missing.

## 2026-04-09 — Apply learnings from user feedback (`6b0f7a9`)

- Processed 2 learning CSVs (23 corrections total).
- Added 5 new classification rules to `siftr.instructions.md`.
- Updated Section 7 to default to processing only the most recent CSV.

## 2026-04-04 — Initial Siftr skill (`850c60b` .. `be78671`)

- Created `siftr/siftr.instructions.md` — 8-step triage workflow.
- Created `modules/Siftr-Inbox.ps1` — Outlook COM backend for categories
  and folder moves.
- Created `siftr/README.md` and `siftr/Siftr_Plan.md`.
- Classification: 7 priority tiers (🔴 through 🔵) with pattern rules.
- Learning mode: low-confidence CSV export and `siftr learn` feedback loop.
- Org context: WorkIQ-based manager/directs/peers with JSON cache.
