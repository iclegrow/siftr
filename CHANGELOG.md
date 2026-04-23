# Siftr Changelog

All notable changes to the Siftr email triage skill are documented here.
This log is maintained by Copilot to preserve context across sessions.

---

## 2026-04-22 — Loop: default final cycle moved to 8 PM

- **Later default stop time**: `siftr loop` now defaults to **8:00 PM local**
  so the final scheduled triage cycle runs at or just after 8 PM before the
  loop stops for the day.
- **Local runner path fixed**: the detached local loop runner now points at the
  standalone `C:\Users\ialegrow\siftr` checkout instead of the removed
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
  clearly asked Ian to confirm ownership of the Xbox follow-up.
- Source: Marjorie Ferris thread (`RE: Package to send`) was incorrectly
  classified 🟢 INFORMED even though the body asked Ian for Thiru's phone
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
- Source: uncategorized earlier replies like `Re: 1:1 Michael/Ian` and
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
