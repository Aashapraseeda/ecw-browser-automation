# Project Documentation — eCW → Patient Forms Now → ReachMyDr Automation

**Audience:** a new engineer taking over this project from scratch, with no prior exposure to the code.
**Companion documents** (same repo root): `MODULE_REFERENCE.md` (Part 5, file-by-file reference for both projects) and `DEBUGGING_HISTORY.md` (Part 8, every major bug hit during development, its root cause, diagnosis, and fix).

This document covers Parts 1, 2, 3, 4, 6, 7, 9, 10, 11, 12 of the full technical handover, written as continuous prose rather than as reference tables, so it reads start to finish like an explanation rather than a spec sheet. The one exception is the architecture diagram, the sequence diagram, and the Nurture-Kids-vs-Lone-Star comparison table, which stay in their original diagram/table form because that structure is the point — the surrounding narrative around each is written in prose.

---

# PART 1 — Project Overview

Two pediatric practices, Nurture Kids Pediatrics and Lone Star Pediatrics, need every well-child visit to result in the right developmental-screening questionnaire — an ASQ, and depending on the child's age, an M-CHAT and/or a TB screening — getting in front of the parent, filled out, and filed permanently into the child's chart. Doing this by hand requires a staff member to pull the next few days' appointments out of the practice's EHR (eClinicalWorks, or "eCW"), work out which of those appointments are well-checks and which forms each child's age calls for, log into a separate product called Patient Forms Now to actually send those forms to the families, log into a third product, ReachMyDr/PCareLink, to text the family a reminder to go fill the form out, come back later to check who has actually completed it, download the resulting PDF, and finally log back into eCW to attach that PDF to the patient's permanent chart documents — all while keeping track, usually just by memory, of who has already been handled so nobody gets double-messaged or missed. None of eCW, Patient Forms Now, or ReachMyDr know the other two exist; there is no API connecting them, and all three are closed, browser-only, third-party portals. This project is a Playwright browser automation that plays the role of that staff member, driving real Chromium sessions against all three real websites end to end, on a repeatable schedule (see Part 9 for the honest current state of that scheduling).

The reason this needed automating in the first place is straightforward: the volume is high and the steps repeat identically for every well-check, every day; memory-based duplicate tracking doesn't scale and getting it wrong means either a family gets messaged twice or, worse, never gets messaged at all; and the actual bottleneck was never any one system being slow, but the friction of handing a patient's identity off between three unconnected systems by hand.

Nurture Kids and Lone Star sit inside the same eCW tenant — they share one eCW login — but they are separate businesses with separate Patient Forms Now accounts and separate ReachMyDr practice configurations. Rather than building one shared codebase for both, two genuinely separate projects exist, `ECW_automation/` for Nurture Kids and `lone_star_automation/` for Lone Star, for one deliberate reason: isolation. A bug or a bad run in one clinic's automation must never be able to touch the other clinic's patients, credentials, or state. Since the two share an eCW login, that isolation has to be enforced explicitly in code — Nurture Kids' automation excludes any appointment at "Lone Star Pediatrics Midlothian" from its own processing, and Lone Star's eCW export applies its own Facility filter at the report-generation step so its export never contains anyone else's appointments in the first place. The full comparison between the two is in Part 6. `ECW_automation/` was built first and is the more battle-tested, reference implementation; `lone_star_automation/` was built afterward, reusing the reference project's proven logic wherever possible and diverging only where Lone Star's actual Patient Forms Now and ReachMyDr accounts genuinely required different handling — discovered through live testing, never assumed up front.

At the architectural level, both projects hang off one shared eCW tenant on one side and one shared ReachMyDr/PCareLink account (covering multiple physical practices, resolved per patient by that patient's own eCW facility) on the other, with each clinic's own pipeline sitting in between as a completely separate Python process, with its own Patient Forms Now account and its own SQLite state database. Nothing is imported across the `ECW_automation/` and `lone_star_automation/` boundary at runtime — they are two independent programs that happen to live in the same repository.

The folder layout reflects that separation directly. At the repo root sit a stale, pre-split `.env` and `requirements.txt` (leftovers from a very early single-file prototype, `ecw_automation.py`, since deleted, that neither current project actually reads), a stale `README.md`, and the two cross-clinic launcher scripts, `run_demo_parallel.py` and `run_production_parallel.py`. Everything that matters lives inside the two project folders:

```
ecw-browser-automation/                  <- repo root
├── .env                                  <- orphaned/legacy, NOT used by either real project
├── .gitignore
├── README.md                             <- stale, predates this whole build-out
├── requirements.txt                       <- root-level, also stale
├── run_demo_parallel.py                   <- runs both DEMO pipelines together
├── run_production_parallel.py             <- runs both PRODUCTION pipelines together
│
├── ECW_automation/                        <- NURTURE KIDS (reference clinic)
│   ├── .env                                <- real credentials for this clinic
│   ├── main.py                             <- THE ENTIRE production pipeline, one flat file (~1200 lines)
│   ├── main_1.py                           <- demo pipeline (test patients only)
│   ├── state_db.py                         <- SQLite state tracking
│   ├── ecw_schedule.xlsx                   <- last downloaded raw export (regenerated every run)
│   ├── filtered_schedule.xlsx              <- last ASQ-eligible-only export (regenerated every run)
│   ├── test_patient_schedule.xlsx          <- shared multi-clinic test data (used by main_1.py)
│   ├── data/patients_state.db              <- the actual SQLite database file
│   └── patients_doc/                       <- per-patient folders of downloaded completed-form PDFs
│
└── lone_star_automation/                  <- LONE STAR PEDIATRICS
    ├── .env                                <- real credentials for this clinic
    ├── main.py                             <- production pipeline (orchestrator only, ~100 lines)
    ├── main_demo.py                        <- demo pipeline (test patients only)
    ├── requirements.txt
    ├── config/settings.py                   <- single source of truth for all config/credentials/mappings
    ├── database/state_db.py                 <- SQLite state tracking (same schema as Nurture Kids')
    ├── ecw/
    │   ├── login.py                          <- shared eCW login + resilience helpers
    │   ├── facility_filter.py                <- Lone-Star-only Facility tab filtering on the report screen
    │   ├── schedule_export.py                <- eCW schedule export (Step 0)
    │   └── chart_upload.py                   <- eCW Chart Documents upload (Step 4)
    ├── patient_forms_now/
    │   ├── login.py                           <- Patient Forms Now login
    │   ├── schedule_import.py                 <- uploads the full Excel into PFN
    │   ├── form_sender.py                     <- eligibility determination + form sending (Step 1)
    │   └── form_downloader.py                 <- completed-form detection + download (Step 3)
    ├── pcarelink/messenger.py                 <- ReachMyDr reminder messaging (Step 2)
    ├── utils/
    │   ├── date_utils.py                       <- DOB → age → ASQ bracket logic
    │   └── logger.py                           <- shared logger factory
    ├── data/patients_state.db
    ├── logs/automation.log
    └── patients_doc/
```

Nurture Kids' `main.py` is a single flat file — every step, from login to Excel parsing to form sending to messaging to upload, lives as a top-level function in one module, which was the fastest way to build it originally and still works, though finding "which function does X" requires a text search rather than a folder to browse. Lone Star's package, by contrast, is split by responsibility, with `main.py` reduced to a thin orchestrator that just calls into each module in sequence — a deliberate improvement made the second time around. If you're extending this project going forward, the Lone Star structure is the better pattern to follow; retrofitting Nurture Kids into the same shape is a reasonable future refactor, but it has not been done — it still works today exactly as a single 1200-line file.

Running any of the pipelines is a matter of invoking the right entry point: `python ECW_automation/main.py` runs Nurture Kids in production, sending real forms and real messages; `python ECW_automation/main_1.py` runs its demo pipeline against a fixed set of test patients only; the same pair exists for Lone Star as `python lone_star_automation/main.py` and `python lone_star_automation/main_demo.py`; and from the repo root, `python run_production_parallel.py` or `python run_demo_parallel.py` launches both clinics' corresponding pipeline as two concurrent subprocesses. Every one of these except the demo variants sends real forms and real text messages to real families the moment it runs — there is no dry-run flag anywhere in this codebase.

---

# PART 2 — Complete End-to-End Workflow

This section narrates a full production run from the moment a person types `python main.py` to the moment the process exits, in the order things actually happen. Lone Star and Nurture Kids differ in a handful of specific places, called out inline as they come up; everything else is identical in spirit even where the exact selectors differ.

The run begins the instant Python imports the module. For Lone Star, importing `config.settings` triggers `load_dotenv()`, which reads that project's own `.env` file into the process environment and builds every constant and mapping the rest of the app relies on — credentials, URLs, ASQ brackets, the facility-to-practice map. For Nurture Kids, `main.py` itself calls `load_dotenv()` and defines all of its constants at the top of the file, at import time. Either way, by the time `asyncio.run(main())` actually starts the event loop, every credential and URL the run will need — `ECW_USERNAME`, `ECW_PASSWORD`, `ECW_LOGIN_URL`, `ECW_EBO_HOME_URL`, the Patient Forms Now credentials (`PFN_*` for Lone Star, `PEDIFORMS_*` for Nurture Kids), the PCareLink credentials and message text, and the various file paths (`EXCEL_PATH`, `DOC_FOLDER`, `STATE_DB_PATH`, `LOG_DIR`) — has already been read out of that project's own `.env` file, none of it hardcoded in source.

`main()` then opens the first of five entirely separate browser sessions this run will use, each its own `async with async_playwright()` block that opens, does its work, and closes before the next one begins — there is no long-lived browser object passed between steps. The first session exists purely to pull the appointment schedule out of eCW. Since eCW has no public API, the only way to get that schedule is to drive a real, visible Chromium window through eCW's own web UI as a logged-in user — visible rather than headless deliberately, since this is a flaky enterprise UI that has repeatedly needed to be watched by a human while debugging it. The login itself fills the username, clicks Next, types the password via real keystrokes rather than a programmatic fill (eCW's password field needs genuine key events to behave), and presses Enter; once the home dashboard's own confirmation element appears, the code waits for eCW's "Building your user experience" loading overlay to disappear — and does so with a retrying wait rather than a single check, because that overlay was proven live to sometimes reappear a moment after it was first confirmed gone (see `DEBUGGING_HISTORY.md`, item 4). A "License Alert" popup is dismissed if eCW happens to show one, which isn't every login.

From the authenticated home page, the code clicks through the side-panel menu into Reports — a sequence of clicks that all use Playwright's `force=True`, because a transient backdrop element (`#pnBackDrop`) or similar was found, unpredictably, to sit briefly on top of otherwise perfectly valid navigation elements and block the click Playwright's default safety check would otherwise refuse to perform (`DEBUGGING_HISTORY.md`, item 7). That "Reports" click spawns a genuine new browser popup window, which Playwright captures and which becomes the page the rest of this step operates on; it's navigated into eCW's eBO Reports module and clicked through to the "Encounter Patient Download" report. If eCW is still processing a leftover, queued report execution from an earlier run on this same account, a "Your report is running" modal appears on top of everything and has to be waited out before any further interaction is possible — and because that modal was observed to sometimes appear again later, after an initial check had already passed, the code checks for it twice: once right after opening the report, and again immediately before touching the date pickers.

Setting the date range and, for Lone Star only, the facility, both happen inside a report-parameter iframe whose contents were found, through live diagnosis, to be genuinely unstable for a period after the iframe itself finishes loading — the underlying report-prompt panel keeps re-rendering internally, with images disappearing and reappearing, in a way that neither Playwright's `networkidle` state nor a fixed sleep reliably outlasts. The fix is a polling loop that waits for the iframe's image count to reach at least two (one calendar icon for the start date, one for the end date) and to stay unchanged across three consecutive checks before treating it as settled — an earlier, looser version of this check accepted a stable count of just one image as "done," which silently broke the end-date picker (`DEBUGGING_HISTORY.md`, items 5 and 6). Only for Lone Star, before the dates are touched, a Facility filter step runs first: it opens the Facility tab, searches for the keyword "lone," and — because the search reliably returns two overlapping results, "Lone Star Pediatrics" and "Lone Star Pediatrics Midlothian" — always picks the second one, confirmed correct through live testing rather than by matching exact text. This is the mechanism that gives Lone Star's export a hard guarantee it will never contain another clinic's appointments; Nurture Kids instead downloads everything and excludes Lone Star's facility name in Python afterward, a real and deliberate asymmetry explained fully in Part 6.

The date range itself is computed as tomorrow through three days out — never today — because the business explicitly wants the automation looking ahead at upcoming appointments rather than processing same-day ones; this replaced an earlier seven-day-from-today window (`DEBUGGING_HISTORY.md`, item 10). Each calendar day click is wrapped in its own retry loop with `force=True`, because live diagnosis showed the target day option could be genuinely visible, enabled, and correctly positioned, and still get blocked by something transiently overlapping it (`DEBUGGING_HISTORY.md`, item 5). Once both dates and, for Lone Star, the facility are staged, one shared OK button submits everything at once, and the code then polls every couple of seconds, for up to four minutes, until eCW's "Select a format" button becomes enabled, meaning the report has actually finished generating server-side. Choosing "Excel data" as the output format and clicking it inside a `page.expect_download()` context downloads a real spreadsheet — roughly seventy columns wide, one row per appointment, including patient account number, name, date of birth, visit type, visit reason, appointment date, and facility — and saves it to a fixed path that gets overwritten every run rather than versioned. The browser closes; everything downstream of this point, until form-sending begins, happens in pure Python with no browser involved at all.

Reading that Excel file is where the actual business logic lives, and it happens as one pass over the rows rather than as separate eligibility and form-selection stages, because the two are computed together. Each row is skipped outright if it has no account number, and for Nurture Kids specifically, skipped if its facility is "Lone Star Pediatrics Midlothian" — the Python-side half of the isolation mechanism described above. The Visit Type cell, formatted as `"CODE : Description"`, is parsed for the description half and checked structurally for whether it represents a well-check at all, by testing whether it ends in the literal suffix `" WC"` — a structural test rather than an enumerated list, adopted specifically because well-checks span ages from two weeks to eighteen years and no fixed list of exact labels could realistically cover that whole range in advance (`DEBUGGING_HISTORY.md`, item 18). The child's age in months at the appointment date is computed from date of birth, and that age drives three genuinely independent form-selection rules. ASQ is resolved via an explicit age-bracket lookup — nine, twelve, fifteen, eighteen, twenty-four, thirty, thirty-six, or forty-eight months — in Lone Star's case purely from age, and in Nurture Kids' case from the Visit Type text itself via a dictionary, a real mechanism difference expanded on in Part 6. M-CHAT is added alongside ASQ only for the eighteen- and twenty-four-month brackets, and is deliberately gated on the bracket number or Visit-Type-text key rather than on the resolved form text itself, because the fifteen-month and eighteen-month brackets happen to resolve to the exact same form text and gating on that shared text once caused M-CHAT to wrongly apply to fifteen-month visits too (`DEBUGGING_HISTORY.md`, item 13). TB is entirely independent of the other two — any well-check from twelve months through eighteen years gets a TB form, even a five-year-old well-check with no applicable ASQ form at all, which under the pre-multi-form architecture would previously have been excluded from the pipeline entirely. A row with no applicable form at all is skipped; a row with at least one is turned into a patient dictionary carrying account number, appointment date, names, facility, visit type, and a `forms` list of one or more `{form_name, form_filename}` entries — with the older singular `form_name`/`form_filename` keys still populated too, now as comma- or underscore-joined summaries, purely so the unchanged SQLite schema's single-string columns still hold something meaningful. Nurture Kids additionally writes a second Excel file at this point containing only the eligible rows, which becomes what actually gets imported into Pediforms later — Lone Star instead imports the full, unfiltered export, letting eligibility decide only who gets searched and sent, not what enters Patient Forms Now in the first place, a further real difference discussed in Part 6.

Before anything is sent anywhere, every eligible patient is checked against the SQLite state database with `is_known(acct_no, appointment_date)`, which looks for a row matching this exact patient-visit — not just this patient, since the same child will have many well-checks over their childhood and each one needs to be tracked independently. Only patients with no existing row become `new_patients`; everyone else is silently excluded from the rest of this run, having presumably already been handled by an earlier one. If `new_patients` comes out empty, the run skips straight ahead to checking for completions from prior runs; otherwise, a second browser session opens, this time against Patient Forms Now.

That session logs in — Lone Star's Organization field has no accessible label and has to be filled via its placeholder text, while Nurture Kids' equivalent field genuinely has a real label, a small but real UI difference between the two accounts — imports the schedule Excel (the full export for Lone Star, the pre-filtered one for Nurture Kids) so the appointments become searchable, and then works through `new_patients` one at a time: searching by account number, opening the patient's record, and, for each patient, opening "+ Send a form" exactly once, checking every box that matches a form in that patient's `forms` list, and clicking "Send form" exactly once for the whole combined set. This was confirmed through a live, read-only inspection of the real UI to genuinely support multi-select — checking three boxes at once leaves all three checked, so one open-panel, one-send cycle correctly handles a patient who needs ASQ, M-CHAT, and TB together. Lone Star's code only carries forward the subset of `new_patients` who were actually, verifiably sent at least one form as `sent_patients`; Nurture Kids, by contrast, marks its entire `new_patients` batch as sent regardless of each individual attempt's outcome — a genuine, confirmed behavioral difference between the two projects rather than an intentional design choice made twice. The browser closes, and `sent_patients` (or the whole batch, for Nurture Kids) is written into SQLite via `insert_form_sent()`, giving each of those patient-visits a row with `status = 'form_sent'` — the earliest possible moment after the real-world action succeeded, deliberately, so a crash immediately afterward loses as little as possible.

A third browser session then handles the ReachMyDr/PCareLink reminder text — for Lone Star, gated on `sent_patients` being non-empty at all, and for Nurture Kids, attempted for the whole `new_patients` batch whenever it's non-empty, following from the difference just described. This account is shared across both clinics and covers several distinct physical practices, so before messaging any given patient, the code resolves which practice actually applies to them from their own eCW facility value, via a small, explicitly confirmed dictionary rather than any heuristic; if a facility isn't in that dictionary, the patient is skipped with a logged warning and never guessed at, and the "Filter by Practice" selection is reopened fresh for every single patient, since two different patients processed back to back in the same run can genuinely belong to different practices. A fixed reminder message is sent to whichever family is currently selected, and the browser closes. Nothing about the outcome of this messaging step is recorded in the database at all — a real, acknowledged gap discussed further in Part 4.

At this point the pipeline stops looking only at patients touched during this specific run and instead asks the database for every patient-visit still sitting at `status = 'form_sent'`, across the entire history of the project, since a patient sent a form three days ago in some earlier invocation is exactly as much in need of a completion check as one sent five minutes ago. If that list is empty, the run skips ahead to cleanup; otherwise a fourth browser session opens against Patient Forms Now again, this time to look for completed submissions. Nurture Kids can shortcut this by checking for the word "Completed" directly on the search-result list row before even opening the patient's record; Lone Star cannot, because its list row only ever shows the word "Downloaded," never literally "Completed," so it has to open the record and check the "Sent forms" table on the detail page instead, at the cost of one extra navigation for every patient not yet actually done. For a patient with at least one completed submission, the code downloads every completed submission it finds, not just the first — a genuine historical gap, since the original single-form-era code only ever grabbed one PDF per patient (`DEBUGGING_HISTORY.md`, item 16). Lone Star's mechanism for this is the more elaborate of the two: each submission renders as its own card under a "Submission Exports" section that populates asynchronously, late enough that even Playwright's `networkidle` wait state was found to fire before the section actually finished rendering — a bug that briefly led to a wrong live conclusion that this whole navigation pattern wasn't supported at all, corrected only once the user supplied a screenshot showing real "Export PDF" buttons that plainly contradicted it (`DEBUGGING_HISTORY.md`, item 14). Each card is checked for an "in_progress" marker and skipped if present; otherwise, expanding "View/Edit Responses" reveals a "Template: name vN" line that gets normalized and matched against the patient's expected form names, reconstructed by splitting the comma-joined `form_name` summary back apart — safe because individual form names never contain commas, unlike the underscore-joined filenames, which can't be split back apart safely. Nurture Kids' simpler UI matches each completed row's own text against expected names directly, without any "Template" concept. Either way, a file already sitting on disk is never re-downloaded, thanks to deterministic, form-specific filenames, and a patient's status only advances to `downloaded` once every one of their expected forms has a matching file on disk — a patient with ASQ done but M-CHAT still outstanding stays at `form_sent` and gets checked again on the next run, rather than being prematurely marked done.

The database is then asked for every patient sitting at `status = 'downloaded'`, and if that list is non-empty, a fifth and final browser session logs into eCW once more to attach the downloaded PDFs to each patient's permanent Chart Documents. It searches the patient by name, opens Patient Docs and then Chart Documents, reads the labels of every document already sitting in that patient's chart by parsing a JSON attribute off each existing document link, and compares that list against every file physically present in that patient's local download folder, uploading only the ones genuinely new. This step already, natively, handles an arbitrary number of files per patient — it simply globs the whole folder — so none of the multi-form work described above required any change here at all. Each successful upload (or a patient who had nothing new left to upload) is recorded as `completed` in SQLite, closing the loop for that patient-visit, and a retention-based cleanup pass then deletes any row that's been sitting at `completed` for longer than the configured retention window, so the database doesn't grow forever. The run ends there, and the Python process exits — either with code zero on success, or with a non-zero code if some unhandled exception escaped `main()` along the way, in which case whatever had already been committed to SQLite up to that point is exactly what survives, since every state-database write in this project is its own small, self-contained connect-execute-commit-close cycle rather than part of one long transaction spanning the whole run.

---

# PART 3 — Data Flow

The cleanest way to see how data actually moves through this system is to trace one piece of information — a single appointment — from the moment it exists only inside eCW to the moment it's a completed, filed chart document, and to notice which representation it takes at each hop.

```
eCW (real appointment data, live)
    │  Playwright browser automation, real login
    ▼
Downloaded Excel report  (ecw_schedule.xlsx — ~70 columns, one row per appointment)
    │  openpyxl, read_only
    ▼
Excel Reader / Eligibility Engine
    (read_eligible_patients_from_excel / read_patients_from_excel)
    │  filters: has acct_no? → not excluded facility? (Nurture Kids)
    │           → is Well Check (text-structural test)? → forms_for_well_check(age)
    ▼
In-memory Patient Objects  (dicts: acct_no, appointment_date, names, facility,
                             forms=[{form_name,form_filename}, ...], joined
                             form_name/form_filename summary strings)
    │
    ├──────────────► state_db.is_known() dedup check
    │                        │
    │                new_patients only
    ▼
Patient Forms Now Import  (full Excel [Lone Star] or filtered Excel [Nurture Kids])
    │
    ▼
Form Sending  (search by acct_no → View → check every form's box → Send once)
    │
    ├──────────────► state_db.insert_form_sent()  (status = 'form_sent')
    │
    ▼
ReachMyDr Messaging  (facility → practice resolution → filter → search → send)
    (no state_db write - not tracked)
    │
    ▼
   [TIME PASSES - parent fills out form(s) - could be minutes, days, or never]
    │
    ▼
state_db.get_pending_patients()  (status = 'form_sent', ALL runs/history)
    │
    ▼
Completed-Form Detection & Download  (search → View → find completed
    submissions → match to expected form → download PDF per submission)
    │
    ├──────────────► state_db.mark_downloaded()  (status = 'downloaded',
    │                 ONLY once every expected form's file exists on disk)
    ▼
Local Filesystem  (patients_doc/{Last First}_doc/*.pdf — one file per form,
                    deterministic filenames, e.g. "Alford_Astoria_TB.pdf")
    │
    ▼
state_db.get_patients_needing_upload()  (status = 'downloaded')
    │
    ▼
eCW Chart Upload  (search patient → Patient Docs → Chart Documents →
    diff against existing docs → upload each new file)
    │
    ├──────────────► state_db.mark_completed()  (status = 'completed')
    ▼
State Database  (final resting state for this patient-visit, until
    retention-based cleanup deletes the row)
    │
    ▼
Logs  (stdout + logs/automation.log [Lone Star only])
```

The property worth internalizing about this diagram is that at every stage, the `status` column in SQLite is the single source of truth for what has and hasn't happened to a given patient-visit — the in-memory patient dictionary itself is never carried across browser sessions. Each of the five sessions rebuilds the patients it cares about from scratch, either by re-reading the Excel file again or by re-reading rows back out of SQLite as dictionaries, and each session is its own independent `async with async_playwright()` block that runs to completion and closes before the next one opens. Nothing about this design assumes the whole pipeline runs in one continuous process without interruption; it's built, deliberately, so that it doesn't have to.

---

# PART 4 — Database / State Database

The whole project relies on one plain SQLite database per clinic — no ORM, raw SQL through Python's built-in `sqlite3` module — stored at `data/patients_state.db` inside each project's own folder, with Nurture Kids and Lone Star each keeping an entirely separate file that is never shared between them. The path is configurable through a `STATE_DB_PATH` environment variable but defaults to that fixed location, and the schema is created idempotently on every single connection, so the very first call to any database function on a brand-new machine creates the file and its one table automatically:

```sql
CREATE TABLE IF NOT EXISTS patients (
    acct_no           TEXT NOT NULL,
    appointment_date  TEXT NOT NULL,
    last_name         TEXT,
    first_name        TEXT,
    visit_type        TEXT,
    form_name         TEXT,
    form_filename     TEXT,
    folder_name       TEXT,
    search_name       TEXT,
    status            TEXT NOT NULL DEFAULT 'form_sent',
    form_sent_at      TEXT,
    completed_at      TEXT,
    PRIMARY KEY (acct_no, appointment_date)
);
```

This exact schema is shared verbatim between both projects, even though the two databases themselves are entirely separate files. Most of the columns are self-explanatory bookkeeping — `last_name`/`first_name` for logging and display, `visit_type` for reference, `folder_name` for locating a patient's downloaded PDFs, `search_name` for the exact "Last,First" string typed into eCW's search box — but two are worth calling out specifically. `form_name` and `form_filename` were originally meant to hold a single form's name and filename; rather than changing the schema when multi-form patients were introduced, these columns now hold a comma- or underscore-joined summary of every form sent for that visit, which the download step later splits back apart to reconstruct what was expected. And `form_sent_at`/`completed_at` are UTC timestamps set exactly once each, the second of which is what the retention-cleanup pass measures against.

The primary key is the compound pair `(acct_no, appointment_date)`, deliberately not `acct_no` alone, because a single child will have many well-check visits across their childhood — a nine-month visit, a twelve-month visit, an eighteen-month visit, and so on for years — and keying on the account number by itself would mean the second such visit is silently treated as "already processed" and skipped forever. Keying on the pair instead gives every individual visit its own independent lifecycle through the state machine, which moves through exactly four values. A new patient-visit is inserted at `form_sent` the moment `insert_form_sent()` runs, right after a form is actually sent. It advances to `downloaded` only once `mark_downloaded()` decides every expected form for that visit has a matching file on disk — until then, it simply stays at `form_sent` and gets re-checked on the next run, which is exactly what happens to a patient with ASQ done but M-CHAT still outstanding. It advances to `completed` once the corresponding PDFs have actually been uploaded into eCW's Chart Documents and `mark_completed()` runs. And eventually, once a row has sat at `completed` for longer than the configured retention window — thirty days by default — `cleanup_old_completed()` deletes it outright, so the table doesn't grow without bound. Every query elsewhere in the pipeline is scoped by this same `status` column, which is the actual mechanism that prevents duplicate work: `get_pending_patients()` only ever returns rows still at `form_sent`, so only those get checked for completion; `get_patients_needing_upload()` only ever returns rows at `downloaded`, so only those get uploaded; and a patient already at `completed`, or already deleted, is invisible to both of those queries and will never be resent, rechecked, or reuploaded.

Duplicate handling actually happens at two separate layers. Before anything is ever sent, `is_known(acct_no, appointment_date)` checks whether this exact patient-visit already has any row at all, regardless of its current status — if so, it's excluded from consideration entirely, and no form or message goes out a second time. Later, at download time, filenames are built deterministically from the patient's name and the specific form's name, never from a timestamp or random value, so a file that's already sitting on disk is simply skipped rather than re-fetched, even though the code will re-process the same patient on every subsequent run until they're fully `completed`.

There is deliberately no explicit retry logic for network or UI failures beyond a handful of small, local retry loops already built into specific steps — the calendar-click retry, the loading-overlay recheck, and so on. The real retry mechanism here is architectural rather than procedural: the whole pipeline is built to be safely re-run from scratch as many times as needed, because dedup means a rerun never resends an already-known visit, deterministic filenames mean a rerun never re-downloads an already-downloaded PDF, the upload step's own existing-document diff means a rerun never re-uploads an already-uploaded file, and every query being scoped by status means a rerun always picks up exactly where a prior run — or a prior crash — left off, with nothing lost and nothing duplicated. That's also why a cron-style repeated invocation, discussed in Part 9, is the intended production model rather than one long-running process with its own internal backoff logic.

Reminder messages have a real, acknowledged gap in this design: they're sent once, immediately after a batch of forms is sent successfully within the same run, and there is no column anywhere tracking whether a reminder actually went out. If the messaging step fails after the form-sent row has already been written — a network blip, a selector miss, an unmapped facility — there is no automatic retry of just the reminder on a later run, because by then the patient is already `is_known()` and will never be selected for messaging again. This is worth flagging to whoever picks this project up next as a real candidate improvement, perhaps a `reminder_sent_at` column plus a small query for "anyone form_sent but never reminded."

Uploaded-PDF tracking works differently from everything else in this database, in that it isn't tracked in SQLite at all at the individual-file level. At upload time, the code reads the patient's actual Chart Documents tree directly out of eCW itself, parsing each document link's own `document-object` JSON attribute for its label, and diffs that live list against whatever files are sitting in the patient's local folder — so eCW's own chart is the real source of truth for "has this specific file been uploaded," while SQLite only ever tracks the coarser `downloaded` → `completed` transition.

Crash behavior follows directly from all of this. If an unhandled exception escapes mid-session — a selector that never resolves, say — it propagates out of `main()` and the process exits non-zero; Playwright's own context-manager cleanup generally closes whatever browser was open in the failing step even though the explicit `browser.close()` line is never reached, a behavior confirmed repeatedly through this project's own live debugging rather than assumed. Whatever was already committed to SQLite at the moment of the crash is exactly what persists, since there's no long-lived transaction spanning multiple steps for a crash to leave half-finished, and nothing downstream of the crash point ever ran at all. The next invocation of `main()` starts completely fresh — a new eCW export, a fresh eligibility pass, fresh dedup checks — and any patient whose processing had genuinely completed and been persisted before the crash is safely skipped as already known, while anyone whose write hadn't happened yet gets reprocessed exactly as if the crash had never occurred. That's the whole point of persisting to SQLite as early and as granularly as possible: the smaller the gap between a real-world action succeeding and that success being recorded, the smaller the blast radius of whatever eventually goes wrong.

Taken together, state persistence in this project is intentionally boring: one SQLite file per clinic, one table, four status values, compound-key dedup, and a standing habit of writing to the database as soon as possible after each real-world action succeeds. There's no message queue, no distributed lock, no cross-process coordination beyond the operating-system-level isolation between the two clinics' pipelines — and, per Parts 8 and 9, a discovered need not to run both clinics' eCW logins at the exact same instant.

---

# PART 6 — Nurture Kids vs. Lone Star

The two projects share the same overall shape but differ in several real, confirmed ways, summarized here as a table since that's genuinely the clearest way to compare them side by side — the reasoning behind each difference is explained afterward, in prose.

| Aspect | Nurture Kids (`ECW_automation/`) | Lone Star (`lone_star_automation/`) |
|---|---|---|
| Code structure | One flat `main.py` (~1200 lines) | Modular package (`ecw/`, `patient_forms_now/`, `pcarelink/`, `database/`, `utils/`, `config/`) |
| eCW report facility scoping | Downloads everything, excludes Lone Star's facility name in Python afterward | Applies a real Facility filter tab on the eCW report screen itself |
| ASQ form selection mechanism | Visit-Type text → dict lookup picks both eligibility signal and exact form | DOB-based age → bracket lookup picks the exact form; Visit Type text only gates "is this a well-check at all" |
| Eligibility data source | Excel | Excel (originally tried Patient Forms Now's own table; proven broken) |
| PFN import scope | Imports a pre-filtered Excel (eligible rows only) | Imports the full, unfiltered Excel |
| Multi-form send success tracking | Marks the whole `new_patients` batch sent, regardless of individual outcome | Only marks patients actually sent a form (`sent_patients`) |
| PCareLink message trigger | Always attempts messaging for the whole batch if non-empty | Only attempts messaging if `sent_patients` is non-empty |
| Completed-form status check location | On the search-result list row itself | On the View/detail page's "Sent forms" table |
| Completed-form matching mechanism | Ancestor row text matched against expected form names | "Template: name vN" line inside an expandable card, after an async-render wait |
| Logging | `print()` only, no file log | Real `logging.Logger`, stdout + rotating file |
| Demo pipeline mechanism | `main_1.py`, has a known unresolved week-filter date-rollover bug | `main_demo.py`, deliberately simpler, text-based mechanism |
| ReachMyDr facility→practice mapping | 5 mapped facilities | 1 mapped facility |

The facility-scoping difference exists because both clinics share one eCW tenant, so both need a hard guarantee they never touch each other's appointments — but Lone Star's own Facility search happened to return two overlapping results ("Lone Star Pediatrics" and "Lone Star Pediatrics Midlothian"), which made a report-level filter, with a hardcoded "always pick the second result," the more reliable option there; Nurture Kids' simpler after-the-fact exclusion was sufficient for its purposes, since it only needs to not include one specific other clinic, not scope itself precisely to one facility with an ambiguous-results problem.

The ASQ-selection difference is a real mechanism divergence, not cosmetic: Lone Star's production path originally tried reading eligibility from Patient Forms Now's own imported table, which was proven, through live debugging, to expose a generic patient-status field rather than the true clinical visit type in that specific location — a bug that produced zero eligible patients in production while staff could see fifty or more just by looking at the same account's patient list. That was fixed by switching Lone Star's eligibility source to the Excel export and its ASQ selection to a DOB-based age bracket, which is what it uses today; Nurture Kids' original Visit-Type-text approach was independently verified correct and left alone.

The import-scope and success-tracking differences follow from how each project evolved rather than from any specific bug: Lone Star's design decided early on that eligibility should only decide who gets searched and sent, not what gets uploaded into Patient Forms Now at all, so it imports the full export; Nurture Kids kept the reference project's original behavior of pre-filtering before import. Similarly, Lone Star's stricter `sent_patients`-only gating for both the state database write and the messaging trigger, versus Nurture Kids' whole-batch approach, reflects how each was actually built rather than a deliberate design parity choice — they are genuinely different, confirmed behaviors, not two implementations of the same intended rule.

The completed-form differences are both directly traceable to live UI inspection: Lone Star's "Today's Patients" list row was found to show the word "Downloaded," never literally "Completed," so its detection has to open the patient's record and check the detail page instead, at the cost of one extra navigation for every not-yet-completed patient; and its "Submission Exports" section, with its expandable "Template: name vN" cards, is a genuinely different and more granular DOM structure than Nurture Kids' simpler completed-forms display, discovered only once a screenshot from the user corrected an earlier, wrong live diagnosis that this pattern wasn't supported on Lone Star's account at all (see `DEBUGGING_HISTORY.md`, item 14).

The logging and ReachMyDr-mapping rows are more straightforward: Lone Star was built with a proper rotating file logger from the start, while Nurture Kids has never been retrofitted with one and still relies entirely on `print()`; and the facility-to-practice mapping sizes simply reflect each clinic's own real, confirmed ReachMyDr configuration — five real facilities for Nurture Kids, one for Lone Star — rather than an arbitrary asymmetry.

In one sentence: Lone Star needed additional logic beyond what was ported from the reference project because, while it shares an eCW tenant with Nurture Kids, it has its own, differently structured Patient Forms Now and ReachMyDr accounts, and every assumption carried over from the reference project had to be independently re-verified live against Lone Star's actual UI — several of those assumptions turned out to be wrong, requiring genuine architecture changes rather than simple configuration swaps.

---

# PART 7 — Complete Execution Timeline

The exact order of operations, for either project, looks like this end to end — Playwright browser sessions are marked explicitly, since everything else runs as plain Python with no browser involved:

```
python main.py
    │
    ▼
main()                                    [orchestrator]
    │
    ├──► schedule_export.run() / ecw_export_schedule()      [Browser #1]
    │        │
    │        ├──► ecw_login() / inline login block
    │        │        └──► _wait_for_loading_overlay_gone()
    │        │        └──► dismiss_license_alert()
    │        ├──► navigate to eBO Reports (force=True clicks)
    │        ├──► navigate to Encounter Patient Download
    │        ├──► _wait_for_report_running_modal_gone()
    │        ├──► wait for iframe + stabilize image count
    │        ├──► _wait_for_report_running_modal_gone()  (again)
    │        ├──► click_calendar_option()  x2 (start, end)
    │        ├──► apply_facility_filter()   [Lone Star only]
    │        ├──► click OK
    │        ├──► poll for report ready
    │        └──► download Excel  → EXCEL_PATH
    │        [Browser #1 closes]
    │
    ├──► read_eligible_patients_from_excel() / read_patients_from_excel()
    │        └──► forms_for_well_check() / _forms_for_patient()  (per row)
    │        [pure Python - no browser]
    │
    ├──► state_db.is_known()  (per patient)  → new_patients
    │
    ├──► [if new_patients non-empty]
    │        │
    │        ├──► form_sender.run_from_excel_list() / pediforms_send_forms()  [Browser #2]
    │        │        ├──► pfn_login() / inline login block
    │        │        ├──► import_schedule()  (full Excel [LS] / filtered [NK])
    │        │        └──► per patient: search → View → _send_forms_for_open_patient()
    │        │        [Browser #2 closes]
    │        │
    │        ├──► state_db.insert_form_sent()  (per sent patient)
    │        │
    │        └──► [if sent_patients non-empty]
    │                 pcarelink.messenger.send_messages() / pcarelink_send_messages()  [Browser #3]
    │                     ├──► login to ReachMyDr
    │                     └──► per patient: resolve_practice_for_facility() → filter → search → send
    │                     [Browser #3 closes]
    │
    ├──► state_db.get_pending_patients()  → pending
    │
    ├──► [if pending non-empty]
    │        │
    │        ├──► form_downloader.run() / pediforms_check_and_download()  [Browser #4]
    │        │        ├──► pfn_login() / inline login block
    │        │        └──► per patient: search → View → check Completed →
    │        │                 download every submission → match to expected form
    │        │                 → mark_downloaded() (only if ALL expected captured)
    │        │        [Browser #4 closes]
    │        │
    │        ├──► state_db.get_patients_needing_upload()  → to_upload
    │        │
    │        └──► [if to_upload non-empty]
    │                 chart_upload.run() / ecw_upload_forms()  [Browser #5]
    │                     ├──► ecw_login() / inline login block
    │                     └──► per patient: search → Patient Docs → Chart Documents
    │                              → diff existing docs → upload new files
    │                     [Browser #5 closes]
    │                 └──► state_db.mark_completed()  (per uploaded patient)
    │
    ├──► state_db.cleanup_old_completed()
    │
    ▼
Program ends (asyncio.run() returns)
```

Read narratively rather than as a tree, this comes down to five sequential browser sessions — export, send, message, download, upload — with pure-Python decision-making happening between each pair of them, and every conditional branch in the middle (`if new_patients`, `if sent_patients`, `if pending`, `if to_upload`) existing specifically so that a run with nothing new to do doesn't open a browser session for no reason.

---

# PART 9 — Production Workflow (current, honest state)

It's worth stating plainly, before anything else, what this section is not: there is currently no cron job, no shell wrapper script, no `PAUSED`-file pause mechanism, no Xvfb virtual display, and no dedicated production virtual environment set up anywhere in this project. Every bit of development and testing so far has happened through manual invocation on a Windows machine, inside a visible, non-headless Chromium window, run directly from a terminal.

What actually happens today is that a person runs `python main.py`, or `main_1.py`/`main_demo.py` for testing, directly from within that project's own folder. To run both clinics together, `python run_production_parallel.py` (or its demo counterpart) launches both as separate operating-system subprocesses via `asyncio.create_subprocess_exec`, streaming both processes' output to one console with a clinic-name prefix so they stay distinguishable. Because both clinics' production pipelines log into the same eCW tenant, running them through this parallel launcher was found, through repeated live testing, to occasionally cause intermittent failures — a stuck login here, a report-generation collision there — specifically when both processes hit eCW's login and report-generation flow at nearly the same moment; this is documented fully in `DEBUGGING_HISTORY.md`. The most reliable practice discovered so far, confirmed through direct experience rather than assumed, is to run the two clinics' production pipelines sequentially — letting one fully finish before starting the other — whenever reliability matters more than total wall-clock time. There is no error alerting, no dashboard, and no email or chat notification on failure built into any of this; the only signal that a run failed is a non-zero process exit code and whatever ended up in the terminal, or in Lone Star's case, in `logs/automation.log`.

None of what a real, unattended production deployment would eventually need exists yet, and it's worth naming honestly, as the natural next step for whoever picks this project up, rather than describing as if it already existed. A scheduler — Windows Task Scheduler, since this currently runs on Windows and `cron` itself doesn't apply unless the project later moves to a Linux server — would need to invoke `main.py` for each clinic on some recurring cadence, which the pipeline's own re-run-safe design (see Part 4) already supports without any further change. Because the browsers currently launch visibly rather than headlessly — a deliberate choice throughout development, since several of the resilience fixes documented in Part 8 were only discoverable by watching the browser directly — a genuinely unattended deployment would need either an untested switch to `headless=True`, whose behavior against these same flaky third-party UIs is not guaranteed identical, or a virtual display server if the project moves to Linux. A pause mechanism, such as a sentinel file a human could drop in to temporarily halt scheduled runs without touching the scheduler configuration itself, is a reasonable idea that simply hasn't been built. And a dedicated virtual environment per project, rather than whatever global Python install currently runs things, would be standard practice for production isolation but similarly isn't set up today.

Recovery after a failure, in the meantime, works exactly as described in Part 4: because the pipeline's SQLite-based state design makes it always safe to re-run manually after a failure, with no cleanup step required first, that re-run is the actual, present-day recovery mechanism — a human notices something didn't produce the expected result and runs the script again by hand. And the error-handling code itself already anticipates a scheduler that doesn't exist yet: nearly every per-patient loop in this codebase wraps its work in a try/except that logs and moves on to the next patient rather than aborting the whole batch, and the browser-session-level functions for checking completions specifically log a message saying the failure "will retry on next scheduled cron run" — language that presupposes the cron setup described above, which, again, isn't actually in place; in practice today, "next scheduled cron run" simply means whenever a person next runs the script.

---

# PART 10 — Architecture Diagram (Text)

```
                                   HUMAN OPERATOR
                                         │
                          runs manually (no scheduler yet)
                                         │
                    ┌────────────────────┴────────────────────┐
                    │                                          │
        python ECW_automation/main.py            python lone_star_automation/main.py
         (or run_production_parallel.py orchestrating both)
                    │                                          │
                    ▼                                          ▼
        ┌───────────────────────┐                  ┌────────────────────────┐
        │  NURTURE KIDS PIPELINE  │                  │  LONE STAR PIPELINE      │
        │  (main.py, flat file)   │                  │  (main.py + package)     │
        └───────────┬─────────────┘                  └────────────┬─────────────┘
                    │                                              │
       ┌────────────┼──────────────┐                ┌──────────────┼───────────────┐
       │            │              │                │              │               │
       ▼            ▼              ▼                ▼              ▼               ▼
   [Browser 1]  [Browser 2]   [Browser 3]       [Browser 1]   [Browser 2]     [Browser 3]
     eCW           PFN          ReachMyDr           eCW           PFN           ReachMyDr
   (export)     (send forms)   (reminders)        (export)     (send forms)    (reminders)
       │            │              │                │              │               │
       └─────┬──────┘              │                └──────┬───────┘               │
             │                     │                       │                       │
             ▼                     ▼                       ▼                       ▼
     ecw_schedule.xlsx     (no DB write - gap)      ecw_schedule.xlsx      (no DB write - gap)
     filtered_schedule.xlsx                                                  (facility-filtered
                                                                               at the eCW report
                                                                               level, not just
                                                                               in Python)
             │                                              │
             ▼                                              ▼
   ┌────────────────────┐                        ┌────────────────────┐
   │  state_db.py (SQLite)│                        │  database/state_db.py│
   │  (Nurture Kids' own)  │                        │  (Lone Star's own)    │
   └──────────┬───────────┘                        └──────────┬─────────┘
             │                                                │
             ▼                                                ▼
       [Browser 4]                                       [Browser 4]
      PFN (download completed)                          PFN (download completed)
             │                                                │
             ▼                                                ▼
     patients_doc/*.pdf                             patients_doc/*.pdf
             │                                                │
             ▼                                                ▼
       [Browser 5]                                       [Browser 5]
      eCW (Chart Documents upload)                       eCW (Chart Documents upload)
             │                                                │
             └───────────────┬───────────────────────────────┘
                             ▼
                   state_db status → 'completed'
                             │
                             ▼
                  cleanup_old_completed() (retention)
```

The diagram is deliberately symmetric, because the two pipelines really are structurally identical — same five browser sessions in the same order, same SQLite state machine underneath — with the asymmetry confined to exactly the places called out in Part 6: Nurture Kids filters Lone Star's facility out in Python after the export, while Lone Star filters at the eCW report level before the export ever leaves eCW; and both pipelines share the same acknowledged gap that ReachMyDr messaging never writes anything back to the database.

---

# PART 11 — Sequence Diagram (Text)

This traces one full patient-visit lifecycle, across every system it touches, for either clinic — the mechanics are identical between the two, with the UI-specific differences already covered in Part 2 and Part 6:

```
Automation          eCW              Patient Forms Now       ReachMyDr           File System         Database
    │                 │                      │                    │                    │                 │
    │──login─────────►│                      │                    │                    │                 │
    │◄──home page──────│                      │                    │                    │                 │
    │──navigate to     │                      │                    │                    │                 │
    │  eBO Reports────►│                      │                    │                    │                 │
    │──set dates,      │                      │                    │                    │                 │
    │  facility────────►│                      │                    │                    │                 │
    │◄──report ready────│                      │                    │                    │                 │
    │──download─────────►│                      │                    │                    │                 │
    │◄──Excel file───────│                      │                    │                    │                 │
    │                    │                      │                    │──save file────────►│                 │
    │                    │                      │                    │                    │                 │
    │──[read Excel, compute eligibility - pure Python, no external call]──────────────────►│                 │
    │                    │                      │                    │                    │                 │
    │──is this patient already known?───────────────────────────────────────────────────────────────────────►│
    │◄────────────────────────────────────────────────────────────────────────────────no─────────────────────│
    │                    │                      │                    │                    │                 │
    │──login─────────────┼─────────────────────►│                    │                    │                 │
    │◄──logged in────────┼──────────────────────│                    │                    │                 │
    │──import schedule───┼─────────────────────►│                    │                    │                 │
    │──search patient─────┼─────────────────────►│                    │                    │                 │
    │──open View──────────┼─────────────────────►│                    │                    │                 │
    │──check ASQ+M-CHAT+TB┼─────────────────────►│                    │                    │                 │
    │──Send form (once)───┼─────────────────────►│                    │                    │                 │
    │◄──sent────────────────┼──────────────────────│                    │                    │                 │
    │                    │                      │                    │                    │                 │
    │──insert_form_sent()─┼──────────────────────┼────────────────────┼────────────────────┼────────────────►│
    │                    │                      │                    │                    │                 │
    │──login───────────────┼──────────────────────┼───────────────────►│                    │                 │
    │◄──logged in───────────┼──────────────────────┼────────────────────│                    │                 │
    │──resolve practice from facility (pure Python lookup)──────────────────────────────────────────────────►│
    │──filter by practice───┼──────────────────────┼───────────────────►│                    │                 │
    │──search patient────────┼──────────────────────┼───────────────────►│                    │                 │
    │──send reminder text──────┼──────────────────────┼───────────────────►│                    │                 │
    │◄──sent────────────────────┼──────────────────────┼────────────────────│                    │                 │
    │                    │                      │                    │                    │                 │
    │       ══════════════ TIME PASSES (parent may or may not fill out the form) ══════════════        │
    │                    │                      │                    │                    │                 │
    │──get_pending_patients()───────────────────────────────────────────────────────────────────────────────►│
    │◄──────────────────────────────────────────────────────────────────this patient (status=form_sent)──────│
    │                    │                      │                    │                    │                 │
    │──login─────────────┼─────────────────────►│                    │                    │                 │
    │──search patient─────┼─────────────────────►│                    │                    │                 │
    │──open View──────────┼─────────────────────►│                    │                    │                 │
    │◄──Completed badge? ──┼──────────────────────│                    │                    │                 │
    │──expand submission───┼─────────────────────►│                    │                    │                 │
    │◄──Template name───────┼──────────────────────│                    │                    │                 │
    │──download PDF─────────┼─────────────────────►│                    │                    │                 │
    │◄──file────────────────┼──────────────────────│                    │                    │                 │
    │                    │                      │                    │──save PDF──────────►│                 │
    │──(repeat per completed submission)                                                                    │
    │──mark_downloaded() [only if ALL expected forms captured]─────────────────────────────────────────────►│
    │                    │                      │                    │                    │                 │
    │──get_patients_needing_upload()─────────────────────────────────────────────────────────────────────────►│
    │◄────────────────────────────────────────────────────────────────this patient (status=downloaded)──────│
    │                    │                      │                    │                    │                 │
    │──login─────────────►│                      │                    │                    │                 │
    │──search patient──────►│                      │                    │                    │                 │
    │──open Chart Documents─►│                      │                    │                    │                 │
    │◄──existing doc labels──│                      │                    │                    │                 │
    │                    │                      │                    │                    │──read files─────►│
    │◄──file list────────────┼──────────────────────┼────────────────────┼────────────────────│                 │
    │──upload new file(s)───►│                      │                    │                    │                 │
    │◄──uploaded─────────────│                      │                    │                    │                 │
    │──mark_completed()──────┼──────────────────────┼────────────────────┼────────────────────┼────────────────►│
    │                    │                      │                    │                    │                 │
```

The gap between "reminder sent" and "time passes" is the one place in this whole diagram with no corresponding database write at all — worth remembering, since it's the same gap flagged as a real limitation in Part 4: nothing here would catch a reminder that silently failed to send.

---

# PART 12 — Explain Like I'm a New Developer

You've just been handed this project, and the single sentence worth keeping in your head while you get oriented is that it's three websites that don't talk to each other, bridged by a browser robot, with a spreadsheet-sized SQLite file remembering who's already been handled.

Start reading code at `lone_star_automation/main.py`. At roughly a hundred lines, it's the orchestrator, and reading it top to bottom tells you the entire shape of the pipeline — export, eligibility, dedup, send, message, download, upload, cleanup — without any of the Playwright selector noise getting in the way. Once that shape is in your head, `ECW_automation/main.py` is the same shape with every step's implementation inlined into one file instead of split across modules; don't try to read it top to bottom, jump straight to the function names that match the shape you already know — `ecw_export_schedule`, `read_patients_from_excel`, `pediforms_send_forms`, `pcarelink_send_messages`, `check_and_download_completed`, `ecw_upload_forms`.

The thing most likely to trip you up if you don't know it going in is that this automation drives real third-party websites it doesn't control, and those websites are not always deterministic. A large fraction of this project's code, and essentially all of its debugging history in `DEBUGGING_HISTORY.md`, is defensive handling of eCW's own UI being flaky — loading overlays that reappear after being confirmed gone, invisible backdrop elements intercepting clicks on perfectly valid buttons, a report iframe whose content flickers in and out of existence while rendering, a "report is running" modal that can reappear mid-flow. When you hit a new failure, your first question should be whether this is the code being wrong or eCW being slow and flaky again, and the pattern that tells the two apart is that a real code bug fails at the same line, the same way, every single time, while eCW's own flakiness fails at a different point, in a different way, run to run. Don't assume it's a code bug just because an exception got raised.

Never test a change by running `main.py` — that's production, and it sends real forms and real text messages to real families the moment it runs, with no dry-run flag anywhere in this codebase. Run `main_1.py` for Nurture Kids or `main_demo.py` for Lone Star instead, both of which are scoped to a small, known set of real test patients through explicit account-number allowlists baked in as a safety gate.

If you need to extend what forms get sent to whom, the single function to change is `forms_for_well_check()` in Lone Star's `patient_forms_now/form_sender.py`, or `_forms_for_patient()` in Nurture Kids' `main.py` — both return a list of form dictionaries, and everything downstream of them, sending, filename generation, download matching, uploading, already handles an arbitrary-length list without modification. Adding a fourth form type at some specific age means adding one more independent rule to that one function; you won't need to touch sending, downloading, or uploading at all.

If you're ever asked to add a third clinic, the right move is to copy `lone_star_automation/`'s folder structure wholesale, since it's the better-organized of the two, point its `.env` at the new clinic's credentials, and expect to spend real time live-testing every UI assumption against that clinic's actual Patient Forms Now and ReachMyDr accounts rather than assuming they behave like either existing one. This project's whole history is a lesson in exactly that: the third-party UI you're assuming looks like account one might not, repeated once already for account two.

And when you do hit a live failure worth debugging, the technique this project was built with, and that solved nearly every non-obvious bug in its history, is a small, read-only Playwright script that logs in, navigates to the exact point of failure, and dumps whatever's actually on the page — element counts, visible text, bounding boxes, a screenshot — without ever clicking Send or submitting anything for real. Reach for that before you reach for a guess based on the error message alone.

By the time you've read this whole document, you should be able to explain to another developer why there are two separate clinic codebases instead of one, why a single run opens five separate browser sessions instead of one long-lived one, why state gets persisted incrementally into SQLite rather than all at once at the end, why ASQ, M-CHAT, and TB are computed as three genuinely independent rules rather than one combined one, and — just as importantly — the honest gap between what this pipeline can already do when a person runs it by hand and what a real, unattended production deployment would still need, laid out plainly in Part 9.
