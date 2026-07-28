# Module Reference — Part 5 of the Project Documentation

Companion to `PROJECT_DOCUMENTATION.md` (Parts 1–4, 6–7, 9–12) and `DEBUGGING_HISTORY.md` (Part 8).

This document goes file by file, for both projects, describing: purpose, every function's inputs/outputs, its dependencies, and how it's called by the rest of the system.

---

## PART 5A — `lone_star_automation/` module reference

### `config/settings.py`

**Purpose**: single source of truth for every credential, URL, path, and business-rule mapping used across the whole project. Every other module imports from here rather than reading `os.environ` directly.

- Calls `load_dotenv()` at import time, reading `lone_star_automation/.env`.
- **Credentials/URLs**: `ECW_USERNAME`, `ECW_PASSWORD`, `ECW_LOGIN_URL` (env-backed, defaults to the hardcoded production URL if unset), `ECW_EBO_HOME_URL` (same pattern), `PFN_ORG`, `PFN_EMAIL`, `PFN_PASSWORD`, `PFN_LOGIN_URL`, `PCARELINK_EMAIL`, `PCARELINK_PASSWORD`, `PCARELINK_MESSAGE`.
- **Paths**: `EXCEL_PATH`, `DOC_FOLDER`, `STATE_DB_PATH`, `LOG_DIR` — all `os.getenv`-backed with sensible project-relative defaults.
- **Facility filtering**: `FACILITY_KEYWORD` ("lone"), `FACILITY_NAME` ("Lone Star Pediatrics Midlothian") — used by `ecw/facility_filter.py`.
- **`WELL_CHECK_VISIT_TYPES`**: a `set` of exact Visit Type description strings known to be well-checks — used as a belt-and-suspenders check alongside the structural `" WC"` suffix test.
- **`ASQ_BRACKET_TO_FORM`**: dict of 8 age brackets (9/12/15/18/24/30/36/48 months) → `{"form_name", "form_filename"}`.
- **`MCHAT_ASQ_BRACKETS = {18, 24}`**, **`MCHAT_FORM`**: gates M-CHAT onto exactly the 18- and 24-month brackets.
- **`TB_MIN_AGE_MONTHS = 12`**, **`TB_MAX_AGE_MONTHS = 216`**, **`TB_FORM`**: gates TB onto any Well Check from 12 months through 18 years.
- **`FACILITY_TO_PRACTICE`**: dict, currently `{"lone star pediatrics midlothian": "Lone Star Pediatrics (Midlothian)"}` (normalized lowercase key).
- **`resolve_practice_for_facility(facility)`**: normalizes the input (`.strip().lower()`), looks it up in `FACILITY_TO_PRACTICE`, returns `None` if unmapped — callers must treat `None` as "skip, log a warning," never as "guess."
- **`STATE_RETENTION_DAYS`**: default 30.

**Depended on by**: every other module in the project (`ecw/*`, `patient_forms_now/*`, `pcarelink/messenger.py`, `database/state_db.py`, `main.py`, `main_demo.py`).

---

### `utils/logger.py`

**Purpose**: one shared logger factory so every module logs consistently.

- **`get_logger(name)`**: returns a `logging.Logger` named `name`, configured once (guards against duplicate handlers on repeated calls within the same process) with:
  - A `RotatingFileHandler` writing to `logs/automation.log`, 5MB max size, 3 backups.
  - A `StreamHandler` to stdout.
  - Both use a bare `"%(message)s"` formatter — no timestamps/levels prefixed, a deliberate choice for clean, readable console output.

**Depended on by**: every module that logs (all of them).

---

### `utils/date_utils.py`

**Purpose**: all DOB/age/ASQ-bracket math, isolated from any UI or I/O concerns — pure functions, easily unit-testable in isolation (though no formal test suite exists in this project).

- **`parse_date_flexible(value)`**: accepts a date string in several possible formats (the Excel's DOB/appointment-date columns aren't perfectly consistent) or a native `datetime`/`date` object, returns a `date`. Raises if truly unparseable.
- **`age_in_months(dob, as_of_date)`**: computes whole months between two dates (calendar-aware, not just day-count/30).
- **`match_asq_bracket(age_months)`**: an explicit cascading set of ranges (e.g. 8–10 → 9, 11–13 → 12, 14–16 → 15, 17–20 → 18, 21–27 → 24, 28–33 → 30, 34–41 → 36, 42–53 → 48) returning the bracket number, or `None` if outside all ranges.

**Depended on by**: `patient_forms_now/form_sender.py` (`read_eligible_patients_from_excel`, `forms_for_well_check`).

---

### `database/state_db.py`

**Purpose**: all SQLite access — the only module that touches `data/patients_state.db`.

- **Schema**: `patients` table as described in Part 4 of `PROJECT_DOCUMENTATION.md`; created idempotently on first connection via `CREATE TABLE IF NOT EXISTS`.
- **`is_known(acct_no, appointment_date)`** → `bool`. `SELECT 1 FROM patients WHERE acct_no=? AND appointment_date=?`.
- **`insert_form_sent(patient)`**: `INSERT OR IGNORE`, sets `status='form_sent'`, stamps `form_sent_at`.
- **`get_pending_patients()`** → `list[dict]`, `WHERE status='form_sent'`.
- **`mark_downloaded(acct_no, appointment_date)`**: sets `status='downloaded'`.
- **`get_patients_needing_upload()`** → `list[dict]`, `WHERE status='downloaded'`.
- **`mark_completed(acct_no, appointment_date)`**: sets `status='completed'`, stamps `completed_at`.
- **`cleanup_old_completed(retention_days)`**: deletes rows where `status='completed' AND completed_at <= now - retention_days`; returns the deleted count.
- **`delete_by_acct_no(acct_no)`**: demo/testing-only reset helper — not called anywhere in the production path (`main.py`), only from `main_demo.py`'s `TESTING_RESET_DEMO_STATE` branch.
- **`normalize_date(value)`**: internal helper coercing any date-like input to the canonical ISO string used as the DB key.

**Depended on by**: `main.py`, `main_demo.py` directly (all state transitions happen in the orchestrator, not inside the browser-automation modules themselves — a deliberate separation of "does the UI action" from "records the state").

---

### `ecw/login.py`

**Purpose**: the one shared eCW login routine, used by every module that needs an authenticated eCW `page` (`schedule_export.py`, `chart_upload.py`).

- **`_wait_for_loading_overlay_gone(page, timeout_ms=60000, retries=3)`**: waits for `#load` to reach `state="hidden"`; if the wait itself raises or the overlay is observed to still be present/unstable, retries up to `retries` times rather than accepting a single check as sufficient (see `DEBUGGING_HISTORY.md` for why).
- **`dismiss_license_alert(page)`**: checks for and clicks through eCW's occasional "License Alert" popup — not every login triggers this, so it's a conditional best-effort check, not a hard requirement.
- **`ecw_login(page)`**: navigates to `settings.ECW_LOGIN_URL`, fills username, clicks Next, waits, fills password (keyboard typing), presses Enter, waits for `#jellybean-panelLink33` (home page confirmation), calls `_wait_for_loading_overlay_gone()`, calls `dismiss_license_alert()`, calls `_wait_for_loading_overlay_gone()` **again** (the overlay was proven live to reappear after the first check passed, specifically around the License Alert dismissal step).

**Depended on by**: `ecw/schedule_export.py::run()`, `ecw/chart_upload.py::run()`.

---

### `ecw/facility_filter.py`

**Purpose**: Lone-Star-specific — scopes the eCW schedule export report to only "Lone Star Pediatrics Midlothian" appointments, at the report-generation level (not just filtered later in Python). This is the mechanism that keeps Lone Star's export from ever containing another clinic's patients, since all clinics using this eCW tenant share the same login.

- **`apply_facility_filter(iframe)`**: clicks the "Facility" tab (`force=True` — an empty sibling badge element was found to intercept this click otherwise), expands "Facility Name," types `settings.FACILITY_KEYWORD` into the search box, clicks Search, **polls the results `<select>`'s option count until it stabilizes** (not a fixed sleep — this exact stability-polling pattern is the one later reused/extended for the report iframe's calendar images in `schedule_export.py`), selects the **second** matching result (Lone Star's own account genuinely lists both "Lone Star Pediatrics" and "Lone Star Pediatrics Midlothian" as separate entries — confirmed live, hardcoded as `.nth(1)` rather than exact-text-matched), clicks "Insert," and verifies the facility landed in the Choice list before returning — raising a hard error if it didn't, rather than silently proceeding unfiltered.

**Depended on by**: `ecw/schedule_export.py::run()`, called after date-range selection, before the shared OK/Finish click.

---

### `ecw/schedule_export.py`

**Purpose**: STEP 0 — logs into eCW, navigates to the Encounter Patient Download report, sets the date range and facility filter, waits for generation, downloads the Excel.

- **`_wait_for_report_running_modal_gone(page1, max_checks=90)`**: waits out eCW's "Your report is running" modal (a leftover/queued execution from this same account's own prior report requests, still processing server-side) — checks every 2 seconds up to `max_checks` times; called **twice** in `run()` (once right after opening the report, and again right before setting dates — the modal was observed live to reappear after the first check passed).
- **`click_calendar_option(iframe, day_str, descriptions_to_try, label, retries=3)`**: tries clicking a calendar day option by role+name+description (trying multiple weekday-description variants, e.g. `"Mo"` then `"Mon"`), falling back to a plain name-only match if all description variants fail — all clicks use `force=True`. Wraps the whole attempt in an outer retry loop (up to `retries` times, 2-second gaps) since live testing showed the day option can be genuinely visible/enabled and still get blocked by something transiently overlapping it.
- **`run(window_days=7, start_offset_days=0)`**: the full STEP 0 sequence — launches Chromium (`headless=False`), calls `ecw_login()`, navigates through the side panel → Menu → Reports (all clicks `force=True`, guarding against `#pnBackDrop` and similar transient elements) into a new popup window (`page1`), through eBO Reports → Administrative Reports → Encounter Patient Download, waits out the report-running modal, waits for the report iframe and **polls until its image count is `>= 2` and stable across 3 consecutive checks** (needed because the date-range picker requires two calendar icons — start and end — and a live run showed the count settling at just 1 image, which an earlier, less strict version of this check wrongly treated as "ready"), waits out the report-running modal again, sets both dates via `click_calendar_option()`, calls `facility_filter.apply_facility_filter()`, clicks the shared OK button, polls (up to 120 × 2s) for the "Select a format" button to become enabled, clicks it, clicks "Excel data" inside a `page.expect_download()` context, saves to `settings.EXCEL_PATH`. Production (`main.py`) calls this with `window_days=2, start_offset_days=1` — tomorrow through +3 days; the demo pipeline calls it with the defaults (today through +2 days) matching the original reference-project demo window.

**Depended on by**: `main.py` (production, non-default args), `main_demo.py`'s pipeline (indirectly, per `TESTING_SKIP_ECW_EXPORT`).

---

### `ecw/chart_upload.py`

**Purpose**: STEP 4 — logs into eCW, uploads every not-yet-uploaded file from a patient's `patients_doc/` folder into their Chart Documents.

- **`_go_to_search(page)`**: presses Escape twice (closes any open patient record/dialog), clicks the patient-hub button if present, waits for and clicks `#jellybean-panelLink65` (`force=True`), waits for the "Last Name, First Name" search box to confirm the search screen is ready. Used both at the start of `run()` and as a recovery/reset step after each patient (success or failure) so the next iteration starts from a known-clean state.
- **`run(patients)`**: launches Chromium (`slow_mo=200` — deliberately slowed down for this step, since document upload dialogs were found to need a moment between actions), logs in, goes to search. Per patient: searches by `search_name`, selects the matching result cell, handles an ambiguous "Please select a patient" popup if it appears, opens "Patient Docs" → quick-searches "chart" → clicks "Chart Documents," reads every existing doc's `document-object` JSON attribute to build a lowercase label list, diffs the patient's local files (by filename-minus-extension, lowercased) against that list to get `new_files`, and for each new file: opens the native file chooser (`page.expect_file_chooser()`), sets the file, confirms a "Please select a category" popup if shown, clicks through two confirmation buttons (`button.commonButton:has-text("OK")`, then `#btnOk`). Returns `uploaded_ok` — every patient who either had nothing new to upload, or had everything uploaded successfully. Any per-patient exception is caught, logged, and the loop calls `_go_to_search()` to recover before continuing to the next patient (one patient's failure never aborts the whole batch).

**Depended on by**: `main.py`, called with `state_db.get_patients_needing_upload()`'s result.

---

### `patient_forms_now/login.py`

**Purpose**: shared Patient Forms Now login, used by `form_sender.py` and `form_downloader.py`.

- **`pfn_login(page)`**: navigates to `settings.PFN_LOGIN_URL`, fills the Organization field via `get_by_placeholder("e.g. Lone Star Pediatrics")` (Lone Star's PFN account has **no accessible label** on this field, only a placeholder — a genuine, confirmed UI difference from Nurture Kids' Pediforms account, which does have a real label), fills Email and Password, clicks Sign in, waits for the post-login landing state.

**Depended on by**: `patient_forms_now/form_sender.py`, `patient_forms_now/form_downloader.py`.

---

### `patient_forms_now/schedule_import.py`

**Purpose**: uploads the schedule Excel into Patient Forms Now so patients become searchable there.

- **`import_schedule(page)`**: clicks "Choose File," sets `settings.EXCEL_PATH` (the **full, unfiltered** export — see Part 6 of `PROJECT_DOCUMENTATION.md` for why this differs from Nurture Kids), clicks "Import schedule," waits for network idle, sets a "week" filter on the resulting patient-list dropdown so the just-imported schedule is what's actually showing.

**Depended on by**: `main.py`'s call into the form-sending step (invoked once, before the per-patient send loop begins).

---

### `patient_forms_now/form_sender.py`

**Purpose**: STEP 1 — the eligibility engine (reads the Excel, decides who needs which form(s)) AND the actual form-sending browser automation.

- **`forms_for_well_check(age_months)`** → `list[dict]`: the three independent form-selection rules — ASQ via `settings.ASQ_BRACKET_TO_FORM[match_asq_bracket(age_months)]` if a bracket matches; M-CHAT appended if the matched bracket is in `settings.MCHAT_ASQ_BRACKETS`; TB appended if `TB_MIN_AGE_MONTHS <= age_months <= TB_MAX_AGE_MONTHS`, **independent of whether ASQ matched at all**. Returns an empty list if none apply.
- **`read_eligible_patients_from_excel(excel_path)`** → `list[dict]`: opens the Excel with `openpyxl` (read-only mode), iterates rows, skips rows with no account number, parses `Visit Type`, determines "is this a Well Check?" structurally (text ends in `" WC"`, cross-checked against `settings.WELL_CHECK_VISIT_TYPES`), computes age via `date_utils.age_in_months(dob, appointment_date)`, calls `forms_for_well_check(age)`, skips the row if the result is empty, otherwise builds and appends the full patient dict (`acct_no`, `appointment_date`, names, `facility`, `visit_type`, `forms` list, joined `form_name`/`form_filename` summary strings, `folder_name`, `search_name`).
- **`_send_forms_for_open_patient(page, patient)`**: with a patient's record already open in PFP, clicks "+ Send a form" **once**, then for every form in `patient["forms"]`, checks the matching checkbox by visible label text, then clicks "Send form" **once** for the whole combined set — confirmed live to genuinely support multi-select (checking three boxes leaves all three checked simultaneously). Returns the list of forms actually checked/sent (used by the caller to build `sent_patients`).
- **`search_and_send_from_list(page, patients)`**: per patient, searches by account number in "Today's Patients," clicks View (skips if not found, logging why), calls `_send_forms_for_open_patient()`, clicks "← Back to today's patients." Returns `sent_patients` — only patients where at least one form was actually checked and sent.
- **`run_from_excel_list(patients)`**: the top-level entry point — launches Chromium, `pfn_login()`, `import_schedule()` (full Excel), `search_and_send_from_list()`, closes browser, returns `sent_patients`.
- **`determine_and_send_from_pfn_table()`**: **dead code**, explicitly flagged — calls a function that was since renamed/removed elsewhere in the file, would raise `NameError` if ever invoked. Deliberately left unfixed because it's unreachable from `main.py`/`main_demo.py` (a historical leftover from before the eligibility source was switched from PFN's own table to the Excel).

**Depended on by**: `main.py` (`read_eligible_patients_from_excel` for eligibility, `run_from_excel_list` for sending); `main_demo.py` (`_send_forms_for_open_patient`, via its own demo-specific send wrapper).

---

### `patient_forms_now/form_downloader.py`

**Purpose**: STEP 3 — the most heavily-iterated file in the whole project. Detects completed submissions and downloads every associated PDF, matching each to its expected form name.

- **`SUBMISSION_EXPORTS_RENDER_WAIT_MS = 3000`**: a fixed extra wait after `networkidle`, specifically for the "Submission Exports" section, which renders its content asynchronously — `networkidle` alone was proven live to return before this section had actually populated (see `DEBUGGING_HISTORY.md`).
- **`_normalize(text)`**: lowercases and strips non-alphanumeric characters — used to make "Template: ASQ-18 Months v1" reliably comparable against a patient's expected form name despite spacing/punctuation differences.
- **`check_and_download_completed(page, patients)`**: per patient, searches by account number in "Today's Patients," opens View. **Checks "Completed" status on the View page's own "Sent forms" table** — not the list row (Lone Star's list row only ever shows "Downloaded," never literally "Completed," a confirmed live UI difference from Nurture Kids — see Part 6). If completed, waits `SUBMISSION_EXPORTS_RENDER_WAIT_MS` for the "Submission Exports" section to actually render, then for **every** "Export PDF" button found: walks up the DOM via `xpath=ancestor::div[@class='card'][1]` to get that specific submission's own card (fixing an earlier strict-mode violation from a too-broad `div.card` filter matching multiple ancestors at once), skips the card if it contains "in_progress" text, otherwise expands "View/Edit Responses" (a read-only click) to reveal a `"Template: <name> vN"` line, normalizes and matches it against the patient's `expected_names` (reconstructed via `patient["form_name"].split(",")` — safe because individual form names never contain commas), and downloads with a deterministic, form-specific filename (falling back to a generic `f"completed form {i+1}"` name if no match is found, rather than guessing a wrong specific name). Skips any file that's already on disk (`os.path.exists`) rather than re-downloading. Only calls `state_db.mark_downloaded()` once **every** expected form for that patient has a matching file on disk — a patient with ASQ done but M-CHAT still pending stays untouched (remains `form_sent`) and gets re-checked on the next run.
- **`run(patients)`**: launches Chromium, `pfn_login()`, calls `check_and_download_completed()`, closes browser.

**Depended on by**: `main.py`, called with `state_db.get_pending_patients()`'s result.

---

### `pcarelink/messenger.py`

**Purpose**: STEP 2 — sends a ReachMyDr/PCareLink reminder text to each patient just sent a form, resolving the correct practice per-patient rather than assuming one fixed practice for the whole run.

- **`send_messages(patients)`**: launches Chromium, logs into the one shared ReachMyDr account. Per patient: calls `settings.resolve_practice_for_facility(patient.get("facility"))`; if `None` (unmapped), logs a warning and **skips** that patient entirely (never guesses a practice); if resolved, logs a `[DEBUG]` line showing the facility→practice resolution (explicitly marked in-code as "remove once verified live" — still present, low-risk to leave, not yet cleaned up), clicks "Filter by Practice" **fresh for this specific patient** (since different patients in the same run can resolve to different practices), selects the resolved practice, searches by account number, opens the message drawer, best-effort selects an "Appointment Scheduling" message type, types `settings.PCARELINK_MESSAGE`, clicks Send, closes the drawer. Closes browser at the end.

**Depended on by**: `main.py`, called only `if sent_patients:` (i.e. skipped entirely if the send step sent nothing this run).

---

### `main.py`

**Purpose**: the production orchestrator — the thinnest file in the project (~100 lines), pure sequencing, no browser-automation logic of its own.

- **`main()`**: calls, in order: `schedule_export.run(window_days=2, start_offset_days=1)` → `form_sender.read_eligible_patients_from_excel(settings.EXCEL_PATH)` → filters against `state_db.is_known()` → `if new_patients: form_sender.run_from_excel_list(new_patients)` → `state_db.insert_form_sent()` for each of `sent_patients` (the real subset actually sent, not the whole batch) → `if sent_patients: pcarelink_messenger.send_messages(sent_patients)` → `state_db.get_pending_patients()` → `if pending: form_downloader.run(pending)` → `state_db.get_patients_needing_upload()` → `if to_upload: chart_upload.run(to_upload)` → `state_db.mark_completed()` for each uploaded patient → `state_db.cleanup_old_completed(settings.STATE_RETENTION_DAYS)`.

**Entry point**: `python lone_star_automation/main.py` → `asyncio.run(main())`.

---

### `main_demo.py`

**Purpose**: a safe, test-patient-only pipeline for validating changes without touching real production data or messaging real families.

- `TESTING_SKIP_ECW_EXPORT = True`: skips STEP 0 entirely, working instead from a pre-existing test Excel.
- `TESTING_RESET_DEMO_STATE = True`: calls `state_db.delete_by_acct_no()` for the known demo account numbers before each run, so the demo can be re-run repeatedly without dedup blocking it.
- **`VISIT_TYPE_TO_FORM` / `VISIT_TYPE_TO_FORM_FILENAME`**: its own, **separate, text-based** dicts (not DOB-based) — a deliberately different, simpler mechanism from production's `forms_for_well_check()`, kept because it's a proven, working, low-risk path for demo/testing purposes specifically.
- **`_build_facility_filtered_excel()`**: builds a small Excel scoped to the demo's known test patients.
- **`read_demo_patients_from_excel()`**: reads that Excel, wraps each patient's single demo form in a `"forms"` list (`[{"form_name": ..., "form_filename": ...}]`) so it's structurally compatible with the shared multi-form send mechanics in `form_sender.py`.
- **`send_forms_via_excel_visit_type()`**: the demo's send-loop wrapper, calling `form_sender._send_forms_for_open_patient()` (updated to call this shared function directly, rather than a since-removed single-form-only equivalent).

**Entry point**: `python lone_star_automation/main_demo.py`.

---

## PART 5B — `ECW_automation/` (Nurture Kids) module reference

Nurture Kids' entire production pipeline lives in **one flat file**, `main.py` (~1200 lines). There is no package structure — every "module" below is a section of that same file, described separately for clarity. `state_db.py` is the one separate file in this project (mirroring `lone_star_automation/database/state_db.py`'s schema/functions verbatim).

### `main.py` — constants and configuration (top of file)

- `load_dotenv()` reads `ECW_automation/.env`.
- `ECW_USERNAME`, `ECW_PASSWORD`, `ECW_LOGIN_URL` (env-backed, defaulting to the hardcoded production URL), `ECW_EBO_HOME_URL` (same pattern), Pediforms credentials (`PEDIFORMS_ORG`/`PEDIFORMS_EMAIL`/`PEDIFORMS_PASSWORD`/`PEDIFORMS_LOGIN_URL`), PCareLink credentials/message.
- Path constants: `EXCEL_PATH` (raw export), `FILTERED_EXCEL_PATH` (eligible-only export — unique to this project, see Part 6), `DOC_FOLDER`, `STATE_DB_PATH`.
- **`VISIT_TYPE_TO_FORM`**: dict mapping exact Visit Type text (e.g. `"9 MONTH WC"`) directly to a resolved ASQ form — the mechanism difference from Lone Star's DOB-bracket approach, described fully in Part 6.
- **`MCHAT_TRIGGER_VISIT_TYPES`**: a **set of Visit Type text keys** (not resolved form names) that additionally get M-CHAT — gated this way specifically to avoid the 15-month/18-month collision bug (both brackets historically resolved to the same ASQ form text; gating on the form text would have wrongly given M-CHAT to 15-month visits too).
- **`TB_MIN_AGE_MONTHS`/`TB_MAX_AGE_MONTHS`/`TB_FORM`**: identical concept to Lone Star's.
- **`EXCLUDED_FACILITY_NAMES = {"lone star pediatrics midlothian"}`**: the Python-level exclusion mechanism (contrast with Lone Star's report-level Facility filter — see Part 6).
- **`FACILITY_TO_PRACTICE`**: dict of 5 real Nurture Kids facilities → ReachMyDr practice names, including `"Peds Center of Round Rock PA" → "Ped Center Of Round Rock"` (the user's final confirmed answer to a genuinely ambiguous facility-name case — see `DEBUGGING_HISTORY.md`).
- **`resolve_practice_for_facility(facility)`**: same normalize-and-lookup pattern as Lone Star's.

### `_wait_for_loading_overlay_gone(page)`

Module-level helper, identical purpose/mechanism to Lone Star's `ecw/login.py` version — called twice each inside `ecw_export_schedule()` and `ecw_upload_forms()` (both steps that log into eCW independently, since there's no shared `ecw_login()` helper function in this project — the login block is duplicated inline in each).

### `_forms_for_patient(visit_type_desc, age_months)`

The Nurture Kids equivalent of Lone Star's `forms_for_well_check()`. ASQ is resolved via `VISIT_TYPE_TO_FORM[visit_type_desc]` (text lookup, may return `None` if the visit type isn't a recognized WC type at all); M-CHAT is appended if `visit_type_desc in MCHAT_TRIGGER_VISIT_TYPES`; TB is appended independently if age is in range. Returns the same `list[dict]` shape as Lone Star's function.

### `read_patients_from_excel()`

The Nurture Kids equivalent of `read_eligible_patients_from_excel()`, with two real differences: (1) it excludes any row whose facility is in `EXCLUDED_FACILITY_NAMES` before doing anything else; (2) after building the eligible-patient list, it **also writes `FILTERED_EXCEL_PATH`** — a second Excel file containing only the eligible rows, which is what gets imported into Pediforms later (Lone Star imports the full raw export instead — see Part 6).

### `_wait_for_report_running_modal_gone()`, `click_calendar_option()`

Same purpose and mechanism as their Lone Star counterparts in `ecw/schedule_export.py` — these functions were **ported to** Lone Star from here (Nurture Kids is the original, reference implementation), not the other direction. Kept as separately-maintained, duplicated code in both projects rather than a shared library, consistent with the project's overall "two independent codebases" philosophy (Part 1.3).

### `ecw_export_schedule()`

The Nurture Kids equivalent of `ecw/schedule_export.py::run()` — same sequence (login inline, navigate to Reports, wait out report-running modal, wait for iframe + stabilize image count `>= 2`, wait out modal again, set dates via `click_calendar_option()`, click OK, poll for report ready, download Excel). **No Facility-filter step** — Nurture Kids doesn't scope the report itself; exclusion happens in `read_patients_from_excel()` instead. Production calls this with the same tomorrow-through-+3-days window as Lone Star.

### `pediforms_send_forms()`

The Nurture Kids equivalent of `form_sender.py`'s `run_from_excel_list()`/`search_and_send_from_list()`/`_send_forms_for_open_patient()` combined into one function: inline Pediforms login, imports `FILTERED_EXCEL_PATH` (the pre-filtered file, not the raw export), then per patient: search, View, open "+ Send a form" once, check every applicable form's box, send once. Returns... notably, **the calling code in `main()` marks the *entire* `new_patients` batch `form_sent`, not just a verified-sent subset** — a real, confirmed behavioral difference from Lone Star (see Part 6).

### `pcarelink_send_messages()`

Same per-patient practice-resolution mechanism as Lone Star's `pcarelink/messenger.py::send_messages()` — resolves practice via `resolve_practice_for_facility()`, skips with a logged warning if unmapped, re-selects "Filter by Practice" per patient. Called from `main()` for the whole `new_patients` batch whenever it's non-empty (not gated on individual send success, unlike Lone Star's `sent_patients`-only gating).

### `check_and_download_completed()` / `pediforms_check_and_download()`

The Nurture Kids equivalent of `form_downloader.py`. Structurally **simpler** than Lone Star's version because Nurture Kids' Pediforms UI doesn't have Lone Star's "Submission Exports"/"Template: name vN" card structure — instead, it checks "Completed" status directly on the **search-result list row** (an efficiency short-circuit Lone Star's UI doesn't support — see Part 6), and for completed patients, matches each completed-submission row's own ancestor `tr`/`div` text against the patient's expected form names. Downloads every completed submission (not just the first — this multi-download behavior was a real historical fix, described in `DEBUGGING_HISTORY.md`), skips already-downloaded files, and only calls `state_db.mark_downloaded()` once every expected form is captured on disk — identical completion-gating logic to Lone Star's, despite the different detection mechanism.

### `ecw_upload_forms()` / `_go_to_search()`

Near-identical port target for Lone Star's `ecw/chart_upload.py` — same `_go_to_search()` recovery pattern, same existing-docs-via-`document-object`-JSON diffing, same native file-chooser upload loop, both using `force=True` on `#jellybean-panelLink65`.

### `main()`

The Nurture Kids orchestrator, same overall shape as Lone Star's `main.py::main()`, with the two confirmed real differences already noted: (1) `state_db.insert_form_sent()` is called for the **whole** `new_patients` batch, unconditionally; (2) `pcarelink_send_messages(new_patients)` is always attempted if `new_patients` is non-empty, regardless of individual form-send outcomes.

**Entry point**: `python ECW_automation/main.py` → `asyncio.run(main())`.

### `main_1.py`

The Nurture Kids demo pipeline — conceptually the counterpart to `lone_star_automation/main_demo.py`, using its own test-patient allowlist and Excel-based visit-type-text mechanics rather than the DOB-based production logic. **Known, still-unresolved issue**: a "week filter" date-rollover bug that has, at various points, blocked test-patient chart access depending on the current date — flagged honestly here as an open item, not silently glossed over (see `DEBUGGING_HISTORY.md` and `PROJECT_DOCUMENTATION.md` Part "Ongoing/unresolved").

### `state_db.py`

Verbatim-identical schema and function set to `lone_star_automation/database/state_db.py` (see Part 5A above for the full description) — a **separate database file** (`ECW_automation/data/patients_state.db`), never shared with Lone Star's.

---

## PART 5C — Repo-root orchestrators

### `run_production_parallel.py`

Launches `ECW_automation/main.py` and `lone_star_automation/main.py` as two **separate OS subprocesses** via `asyncio.create_subprocess_exec`, streaming each process's stdout/stderr with a distinguishing prefix (`[reference_clinic_production]` / `[lone_star_production]`). Reports each process's exit code and elapsed time, plus total wall-clock time at the end. Includes `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` to prevent Windows console crashes on non-ASCII characters (em-dashes) in log output. See Part 9 of `PROJECT_DOCUMENTATION.md` for the real, observed reliability caveat with running both clinics' eCW logins simultaneously this way.

### `run_demo_parallel.py`

Same mechanism, targeting `ECW_automation/main_1.py` and `lone_star_automation/main_demo.py` instead.
