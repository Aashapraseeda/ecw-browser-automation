# Debugging History — Part 8 of the Project Documentation

Companion to `PROJECT_DOCUMENTATION.md` (Parts 1–4, 6–7, 9–12) and `MODULE_REFERENCE.md` (Part 5).

Every entry follows the same shape: **Symptom → Diagnosis → Why earlier attempts failed (if any) → Final fix → Where it lives in code.** These are written in the order they were actually encountered and solved during this project's development, not grouped by theme, since later bugs sometimes reused patterns discovered in earlier ones — that lineage is itself useful context for a new maintainer.

---

## 1. Lone Star production found 0 eligible patients, staff manually saw 50+

**Symptom**: Lone Star's production pipeline consistently reported zero eligible patients for form-sending, even on days staff could see 50+ appropriate well-check appointments by looking directly at Patient Forms Now's own patient list.

**Diagnosis**: the original eligibility mechanism read a "Visit type" column *inside Patient Forms Now's own imported table*, not eCW's Excel export. Live inspection showed this PFN column is a **generic patient-status field** (things like "New," "Active"), never the actual clinical visit type ("9 MONTH WC," etc.) — so the eligibility filter was matching against the wrong data entirely, structurally guaranteed to return close to nothing.

**Why this wasn't a "fix the matching logic" bug**: no amount of adjusting the text-matching rules could work, because the correct information (real clinical visit type) was never present in the field being read at all — it lived only in the original eCW Excel export, which the old code exported but never consulted directly for eligibility.

**Final fix**: rewrote eligibility to read directly from the eCW Excel export (`read_eligible_patients_from_excel()` in `patient_forms_now/form_sender.py`), reusing Nurture Kids' proven Visit-Type/DOB-based pattern, while keeping the DOB-based age/form-selection logic intact and importing the full Excel into PFN unchanged (only the *eligibility source*, not the *import content*, changed).

**Code**: `lone_star_automation/patient_forms_now/form_sender.py::read_eligible_patients_from_excel()`, `config/settings.py::ASQ_BRACKET_TO_FORM`/`utils/date_utils.py`. The old, now-dead `determine_and_send_from_pfn_table()` function is left in place, explicitly flagged as unreachable, rather than silently deleted.

---

## 2. Session conflicts / Chrome profile / Playwright context isolation

**Symptom**: concern (and, later, live evidence) that running both clinics' pipelines "at once" could cause one automation to interfere with the other's eCW session, since both log into the **same eCW tenant/account**.

**Diagnosis, first pass**: each `run()` function opens its own `async with async_playwright()` block and its own `browser.new_context()` — Playwright contexts are isolated by design (separate cookies, separate storage), so *in principle* two simultaneous Chromium instances logging into the same eCW account as two independent browser sessions should behave like two different people logging in from two different computers — which eCW is generally built to tolerate.

**Why this theory wasn't sufficient in practice**: repeated live runs via `run_production_parallel.py` showed real failures that moved around between runs — a password field disabled on one run, a connection timeout on another, a stuck calendar click on a third — with no single, consistent failure point. That pattern (moving failure location, not a fixed one) is the signature of **resource/timing contention** rather than a deterministic code bug: two processes hitting eCW's login and report-generation endpoints at nearly the same instant appear to sometimes trip up eCW's own session/report-queue handling for that account, even though each Playwright context is technically isolated on the client side.

**Final fix**: **no code fix was implemented** for this — it remains a real, open architectural question. The practical mitigation communicated and used going forward was: **run the two clinics' production pipelines sequentially, not via the parallel launcher, when reliability matters more than wall-clock time.** This is documented plainly, not glossed over, in Part 9 of `PROJECT_DOCUMENTATION.md`. A more thorough investigation (e.g. deliberately staggering subprocess launch by N seconds inside `run_production_parallel.py`, or testing whether eCW itself rate-limits/queues concurrent report requests from one account) is a reasonable next step for whoever picks this up.

---

## 3. eCW connection timeout under Playwright, but manual login worked fine

**Symptom**: `net::ERR_CONNECTION_TIMED_OUT` when Playwright navigated to the eCW login URL — but the same URL, opened manually in a normal browser on the same machine at the same time, loaded and logged in without any issue.

**Diagnosis**: since manual access worked, this ruled out the eCW server itself being down, and ruled out a genuine network/DNS/firewall problem for that URL in general. The question became whether something Playwright/Chromium-specific (proxy inheritance, a stale/broken automation-launched browser profile, a transient one-off hiccup) was responsible.

**Investigation approach**: rather than guessing, verified directly — confirmed the URL was reachable via `curl` from the same shell the automation would run in, then ran a fresh, isolated Playwright test navigating to just that URL with no other automation logic involved.

**Finding**: the fresh, isolated test succeeded — the failure did not reproduce in isolation. This pointed to the earlier failure being a **transient condition** (a one-off slow response, a brief real network hiccup, or contention with another already-running automation process — see item 2), not a persistent Playwright-vs-manual-browser configuration difference.

**Final fix**: no permanent code change; the resolution was diagnostic confirmation that the issue was transient/environmental rather than a bug requiring a fix, communicated honestly to the user rather than inventing a speculative "fix" for a problem that wasn't reproducible.

---

## 4. `#load` loading-overlay bug (eCW home page)

**Symptom**: automation would proceed past eCW's home-page load, only to fail moments later on the very next click — as if the page wasn't actually fully ready despite the login sequence completing.

**Diagnosis**: eCW shows a "Building your user experience" loading overlay (`#load`) after login. A single check-once-and-proceed wait for this element to become hidden was **not sufficient** — live testing showed the overlay could reappear briefly (e.g. around the License Alert dismissal step) even after an initial hidden-state check had already passed.

**Why the first fix attempt (a single `wait_for(state="hidden")`) failed**: it only guaranteed the overlay was hidden *at the moment checked* — it made no guarantee about what happened a second later, and the overlay was observed to come back.

**Final fix**: `_wait_for_loading_overlay_gone(page, timeout_ms=60000, retries=3)` — re-checks up to 3 times rather than trusting a single pass, and is called **twice** inside the login sequence (once right after the home page loads, and again immediately before returning control to the caller, since the overlay was specifically observed to reappear around the License Alert step in between those two points).

**Code**: `lone_star_automation/ecw/login.py`; the equivalent module-level function in `ECW_automation/main.py`.

---

## 5. Calendar day click failing despite the element being visible and enabled

**Symptom**: `click_calendar_option()` would time out trying to click a specific day in the date-range picker, even though a live diagnostic confirmed the target day option was genuinely visible, enabled, and had a real, non-zero bounding box.

**Diagnosis, first theory (wrong)**: assumed the day option itself was missing or mis-labeled — checked description-variant matching (`"Mo"` vs `"Mon"`), added a plain-name fallback. This helped in some cases but did not fully resolve the failures.

**Diagnosis, second theory (correct, confirmed via live inspection)**: the report iframe's own content was found to be **flickering** — a diagnostic script that dumped the iframe's image count immediately after "the iframe loaded" found **zero images present**, moments after the load event fired. The report's prompt-tab panel (Additional Prompts / Facility / Provider / Payer / Patient / Others) was still internally re-rendering; the calendar icon (and thus its resulting day-picker popup) hadn't actually stabilized yet, even though Playwright's own load-state signals said it had.

**Why `networkidle`/a fixed wait wasn't enough**: `networkidle` reflects network activity, not client-side re-rendering — this UI's internal re-render cycle produces no new network requests, so `networkidle` was satisfied while the DOM was still actively changing underneath it.

**Final fix, in two layers**:
1. A **stability-polling loop**: poll the iframe's image count every second, requiring it to be non-zero **and unchanged across 3 consecutive checks** before proceeding — not just "eventually non-zero."
2. `force=True` on the actual day-option click, plus an outer retry loop (up to 3 attempts) — since even after the iframe stabilizes, the click can still be transiently blocked by something overlapping it momentarily (the same category of issue as item 7 below, just applied to a different element).

**Code**: `lone_star_automation/ecw/schedule_export.py::run()` (the stability-polling block) and `::click_calendar_option()` (the retry+force block); mirrored in `ECW_automation/main.py`.

---

## 6. Stability check accepted "1 image" as done when 2 were actually needed

**Symptom**: even after adding the stability-polling loop from item 5, one live run still failed, later in the flow, on the **end**-date calendar click specifically (`iframe.get_by_role("img").nth(1)`).

**Diagnosis**: the date-range picker requires **two** calendar icons — one for the start date, one for the end date. The stability check's original condition was "count is non-zero and unchanged across 3 checks" — it did not require any specific *minimum* count, so a run where the count stabilized at exactly **1** image (the start-date icon rendered, the end-date icon not yet) was wrongly accepted as "stable and ready."

**Why this wasn't caught by item 5's fix immediately**: the earlier fix correctly addressed *instability* (images appearing/disappearing) but hadn't yet addressed *incompleteness* (a stable-but-wrong count) — two different failure modes that happened to look similar in the logs.

**Final fix**: tightened the condition to require `count >= 2` specifically, not just "non-zero," before considering the wait satisfied — reusing the same 3-consecutive-checks stability requirement on top of that minimum. Extended the max wait from 30 to 90 seconds at the same time, since the true 2-image state was observed to sometimes take longer to arrive than the original window allowed.

**Code**: same location as item 5, `ecw/schedule_export.py::run()`'s polling loop (`count >= 2 and count == previous_count`).

---

## 7. `#pnBackDrop` and other transient elements intercepting navigation clicks

**Symptom**: clicks on real, valid navigation elements (the side-panel menu icon, "Menu," "Reports," the Facility tab) intermittently failed with Playwright's actionability-check error, reporting that a *different* element (varying between runs — sometimes `#pnBackDrop`, sometimes an empty sibling badge element) was intercepting the click at that coordinate.

**Diagnosis**: eCW's own UI creates short-lived overlay/backdrop elements during its internal transitions (opening a menu, switching tabs) that briefly sit on top of otherwise-normal, already-rendered, clickable elements. Playwright's default actionability check (which refuses to click an element with something else on top of it, to avoid clicking the wrong thing) was — correctly, by design — refusing these clicks. The underlying navigation elements were never actually broken; they were just transiently covered.

**Why a plain retry-without-`force` wasn't the fix**: waiting and retrying the *same* click could still lose the race against the next transient overlay appearing, since these overlays are unpredictable in timing and don't correspond to any single stable "wait for X to disappear" condition worth targeting individually for every different overlay element.

**Final fix**: `force=True` on the specific clicks proven to be affected — the side-panel menu icon (`#jellybean-panelLink4`), "Menu," "Reports," the Facility tab click (`facility_filter.py`), the calendar day clicks (item 5), and `#jellybean-panelLink65` in both `chart_upload.py` and Nurture Kids' equivalent search-navigation code. `force=True` bypasses Playwright's overlap/actionability check and clicks the element's coordinates directly — this is safe specifically *because* the underlying element was independently confirmed (via live diagnostics) to be the correct, real, intended target; `force=True` is not a blanket "ignore all problems" fix and was never applied speculatively to elements that hadn't been confirmed present and correct first.

**Code**: `lone_star_automation/ecw/schedule_export.py`, `ecw/chart_upload.py`, `ecw/facility_filter.py`; mirrored equivalents in `ECW_automation/main.py`.

---

## 8. "Your report is running" modal reappearing after already being checked

**Symptom**: even after confirming (via `_wait_for_report_running_modal_gone()`) that no report-running modal was blocking the screen, a later click on the date picker would still fail — as though a modal had appeared *after* the earlier check had already passed.

**Diagnosis**: this modal reflects a **leftover/queued report execution from that same eCW account's own prior report requests**, still processing server-side — it can appear (or reappear) at any point while the report-parameter screen is open, not just once at the start, since it's driven by server-side report-queue state rather than anything the client-side automation controls directly.

**Why checking it once at the start wasn't sufficient**: the modal's appearance is asynchronous relative to the automation's own navigation steps — passing the check once says nothing about whether a *new* report execution (possibly from a different, unrelated prior run of this same account) might start queuing and surface the modal moments later.

**Final fix**: call `_wait_for_report_running_modal_gone()` **twice** — once right after opening the report, and again immediately before interacting with the date pickers — rather than assuming a single early check covers the whole subsequent flow.

**Code**: `lone_star_automation/ecw/schedule_export.py::run()`; the equivalent `_wait_for_report_running_modal_gone()` in `ECW_automation/main.py`.

---

## 9. Facility filtering — Lone Star vs. Nurture Kids need genuinely different mechanisms

**Symptom/requirement**: both clinics share one eCW tenant, so both need a guarantee their automation never touches the other clinic's appointments.

**Diagnosis**: two different, valid mechanisms were available — filter at the eCW **report-generation** level (scope the export itself), or filter **after export**, in Python, by excluding a known facility name from the parsed rows. Neither is objectively "more correct" in isolation; the choice mattered because of a subtlety in Lone Star's own eCW account: the live Facility results list, once searched, returned **two** matching entries ("Lone Star Pediatrics" and "Lone Star Pediatrics Midlothian") rather than one exact match — meaning naive exact-text selection would have been ambiguous or wrong depending on result ordering.

**Final fix**: Lone Star applies a real report-level Facility filter (`ecw/facility_filter.py`), confirmed via live testing to require selecting the **second** of the two results (hardcoded `.nth(1)`, not text-matched, since both results share overlapping text) — this gives Lone Star's export a hard guarantee it never contains another clinic's appointments in the first place. Nurture Kids instead excludes Lone Star's facility name in Python after downloading the full export (`EXCLUDED_FACILITY_NAMES`) — sufficient for its purposes (it only needs to *not include* Lone Star, not scope itself to one specific facility with an ambiguous-results problem).

**Code**: `lone_star_automation/ecw/facility_filter.py::apply_facility_filter()`; `ECW_automation/main.py`'s `EXCLUDED_FACILITY_NAMES` check inside `read_patients_from_excel()`.

---

## 10. Date range change: 7-day-from-today → tomorrow through +3 days

**Symptom/requirement**: an explicit, deliberate business-logic change — the export window needed to start tomorrow, not today, and span 3 days total (today+1 through today+3), not a full week from today.

**Diagnosis**: this was not a bug, but a scoping change — the original 7-day-from-today window was intentionally too broad for current needs, and included *today's* appointments, which the business wanted excluded (the automation should look forward, not process same-day appointments).

**Final fix**: `schedule_export.run(window_days, start_offset_days)` was parameterized (it previously had a fixed default), and production's call site was changed to `run(window_days=2, start_offset_days=1)` — `start_date = today + 1`, `end_date = start_date + 2`, giving exactly today+1/+2/+3. The demo pipeline's call site was **deliberately left unchanged** (`start_offset_days=0`, the original today-through-+2 window), since the demo's purpose is stable, repeatable testing against known test patients, not mirroring the production date logic. The same parameterization and call-site change was applied to Nurture Kids' `ecw_export_schedule()`.

**Code**: `lone_star_automation/ecw/schedule_export.py::run()` signature and its call in `main.py`; `ECW_automation/main.py::ecw_export_schedule()`'s equivalent call in `main()`.

---

## 11. ReachMyDr messaging sent from one hardcoded practice, regardless of the patient's actual facility

**Symptom/requirement**: the ReachMyDr/PCareLink account used for messaging covers **multiple** real practices, but the original messaging code selected one fixed, hardcoded practice for the entire run — meaning a patient from a different facility than the hardcoded one would get a reminder filed under (and possibly visible/attributed to) the wrong practice.

**Diagnosis**: needed a genuine per-patient facility → practice mapping, built from real, confirmed data — not inferred or guessed.

**A real detour worth documenting**: an early round produced one confirmed mapping for an ambiguous case ("Peds Center of Round Rock PA" → "Pediatric Center Of Round Rock"), which the user later explicitly asked to **discard entirely**, instructing that the mapping be re-derived from actual data with no carried-over assumption, and that any *remaining* ambiguous cases be asked about individually rather than guessed. The second, clean pass through the same ambiguous facility name resulted in the **opposite** answer ("Ped Center Of Round Rock") — a useful reminder that a first confirmed answer under time pressure isn't necessarily durable, and that re-deriving from scratch with a clear head can produce a different, more considered result.

**Final fix**: `FACILITY_TO_PRACTICE` dict (normalized-lowercase keys) + `resolve_practice_for_facility()` in both projects' settings; `pcarelink/messenger.py`/`pcarelink_send_messages()` resolve the practice fresh per patient (not once per run), log a warning and **skip** (never guess) any patient whose facility isn't in the map, and re-open "Filter by Practice" per patient since different patients in the same run can resolve to different practices.

**Code**: `lone_star_automation/config/settings.py::FACILITY_TO_PRACTICE`/`resolve_practice_for_facility()`, `pcarelink/messenger.py::send_messages()`; `ECW_automation/main.py`'s equivalent dict/function/`pcarelink_send_messages()`.

---

## 12. Live PCareLink `send_messages()` verification blocked by Claude Code's own safety system

**Symptom**: an attempt to run a live verification of the messaging step's actual send behavior was blocked — not by any error in the code or the target website, but by Claude Code's own auto-mode safety classifier, which does not permit autonomously executing an action that sends a real message to a real person without more explicit, fresh authorization for that specific action.

**Diagnosis**: this is a deliberate safety boundary, not a bug to route around. Live verification of an irreversible, externally-visible action (an actual text message to a real family) is exactly the category of action that requires the human operator's own explicit, in-the-moment confirmation — not something that should be silently bypassed or worked around by an automation change.

**Resolution**: the block was respected as-is. The user was told plainly what happened and why, given the exact command to run themselves if they wanted to perform that specific verification, and told they were welcome to give fresh, specific confirmation for that exact action if they wanted it done directly. No code or process change was made to bypass this.

---

## 13. M-CHAT incorrectly triggered for 15-month visits (both projects)

**Symptom**: caught during offline verification, before ever reaching the user — M-CHAT was being added for some 15-month-old patients, who should only get ASQ.

**Diagnosis**: the 15-month and 18-month ASQ brackets both resolve to the **same underlying form text** ("ASQ-18/18 Months" in Lone Star's bracket mapping; the equivalent collision in Nurture Kids' `VISIT_TYPE_TO_FORM` dict). The original M-CHAT gating logic checked "does this patient's *resolved form* match the ASQ-18-months form?" — which is true for **both** the 15-month and the true 18-month bracket, since they share that resolved text.

**Why this is a subtle bug, not an obvious one**: the bug is invisible if you only look at the *output form name* — both brackets genuinely do get an "ASQ-18 Months" form, correctly. It only shows up if you specifically check *which age bracket* a given form assignment came from.

**Final fix**: gate M-CHAT on the **bracket number itself** (Lone Star: `bracket in MCHAT_ASQ_BRACKETS = {18, 24}`) or the **Visit Type text key** (Nurture Kids: `visit_type_desc in MCHAT_TRIGGER_VISIT_TYPES`) — i.e., on the *input* that determined the form, never on the *resolved output form name* itself, since the output name is exactly the thing proven to collide across distinct brackets.

**Code**: `lone_star_automation/config/settings.py::MCHAT_ASQ_BRACKETS`, `patient_forms_now/form_sender.py::forms_for_well_check()`; `ECW_automation/main.py::MCHAT_TRIGGER_VISIT_TYPES`, `_forms_for_patient()`.

---

## 14. "Submission Exports (not supported)" — a wrong live diagnosis, corrected by the user's screenshot

**Symptom**: an early live diagnostic concluded that Lone Star's Patient Forms Now "Today's Patients → View → Export PDF" flow did not work — the "Submission Exports" section appeared to say "No submissions linked," suggesting this navigation pattern (used successfully by Nurture Kids) simply wasn't supported on Lone Star's account.

**Why this diagnosis was wrong**: the diagnostic script read the page's content immediately after Playwright's `networkidle` state was reached. `networkidle` reflects **network** activity quieting down — it does not guarantee that **client-side, JavaScript-driven** rendering has finished painting. The "Submission Exports" section populates asynchronously, after `networkidle` fires, so the diagnostic was reading a real page state — just one that hadn't finished rendering yet.

**How it was caught**: the user provided a screenshot of the same page, taken a bit later / after manual interaction, showing genuinely real "Export PDF" buttons and "(in_progress)" entries — directly contradicting the "not supported" conclusion. This was accepted immediately as correct evidence over the earlier automated conclusion, without defensiveness — the mistake was explicitly acknowledged, the exact mechanism explained (premature read racing an async render), and the conclusion re-verified live with a longer wait before proceeding.

**Final fix**: `SUBMISSION_EXPORTS_RENDER_WAIT_MS = 3000` — an explicit additional fixed wait, specifically for this section, layered on top of (not instead of) the existing `networkidle` wait. This led directly to the full migration of `form_downloader.py` to use the "Today's Patients → View → Export PDF" pattern (matching Nurture Kids' approach) with a more sophisticated card/`Template:`-name-matching mechanism suited to Lone Star's actual DOM structure (see Module Reference, `form_downloader.py`).

**Code**: `lone_star_automation/patient_forms_now/form_downloader.py::SUBMISSION_EXPORTS_RENDER_WAIT_MS`, `check_and_download_completed()`.

---

## 15. Off-by-one bug in a diagnostic script's "expand all submissions" loop

**Symptom**: a one-off diagnostic script written to expand every "View/Edit Responses" button on a page and read their contents exited early, missing some submissions.

**Diagnosis**: the loop iterated over a live `count()` of matching buttons — but expanding one "View/Edit Responses" button changed the page's own DOM (its button might disappear or change state once expanded), which shifted the remaining count out from under the loop mid-iteration, causing an early exit before all submissions had actually been read.

**Final fix**: caught and corrected within the same diagnostic session — re-written to snapshot the button locators once up front rather than re-querying a live, changing count on each iteration. Explicitly explained to the user as a bug in the *test script itself*, not the production code, since this diagnostic tool never shipped as part of `form_downloader.py`.

---

## 16. Multi-PDF download — only the first completed form was ever downloaded

**Symptom/requirement**: with the introduction of multi-form patients (ASQ + M-CHAT + TB), the download step needed to fetch **every** completed submission for a patient, not just one — the original single-form-era code only ever grabbed the first "Export PDF" it found.

**Diagnosis**: this wasn't so much a "bug" as an assumption baked into the original single-form design that needed to be revisited once patients could have 1, 2, or 3 forms outstanding at once. Naively looping over every "Export PDF" button also surfaced a real Playwright strict-mode violation — an earlier attempt to find each submission's containing card via a broad `div.card` filter matched **multiple ancestor divs at once** for a given button (nested `.card` elements), which Playwright's strict mode correctly rejected as ambiguous.

**Final fix**: `xpath=ancestor::div[@class='card'][1]` — walks up to the **nearest** ancestor with exactly that class, resolving the ambiguity deterministically; each matched card's own "Template: name vN" text (after expanding "View/Edit Responses") is used to assign that specific PDF a specific, deterministic filename; a patient is only marked `mark_downloaded()` once every expected form name has a matching file on disk (see Part 4 of `PROJECT_DOCUMENTATION.md`). Verified offline via a custom mock harness covering 1-form, 2-form, and 3-form scenarios, plus a partial-completion-across-multiple-runs scenario (some forms done today, others days later) — confirmed no duplicate downloads and no overwrites in any case.

**Code**: `lone_star_automation/patient_forms_now/form_downloader.py::check_and_download_completed()`; the analogous multi-download loop in `ECW_automation/main.py::check_and_download_completed()` (simpler mechanism, same completion-gating logic — see Part 6 for the mechanism difference).

---

## 17. Excel-based eligibility rewrite — see item 1

Covered fully above; listed again here only because the user's requested documentation structure calls it out as its own named historical bug distinct from the general "0 eligible patients" symptom. No additional detail beyond item 1.

---

## 18. Well Visit filtering — moving from enumerated lists to a structural check

**Symptom/requirement**: TB screening needed to apply to **any** Well Check from 12 months through 18 years — a range far too wide to enumerate every exact Visit Type label in advance (unlike the original ASQ-only logic, which only needed to recognize a small, fixed set of well-check ages).

**Diagnosis**: the existing "is this a Well Check?" logic (in both projects, originally) relied on matching against known, enumerated Visit Type strings or a small lookup dict — workable when only 8 specific ASQ ages mattered, but incomplete once *any* age from 1 to 18 years needed to be correctly identified as a Well Check for TB purposes.

**Final fix**: switched the core "is Well Check" test to a **structural** one — does the parsed Visit Type text end in `" WC"` — which generalizes correctly across the full age range without needing to enumerate every specific label. Both projects retain their original enumerated set/dict (`WELL_CHECK_VISIT_TYPES` in Lone Star, the dict-membership check in Nurture Kids) as an **additional**, not replacement, cross-check, and Lone Star's fallback also checks the `Visit Reason` column as a further belt-and-suspenders signal.

**Code**: `lone_star_automation/patient_forms_now/form_sender.py::read_eligible_patients_from_excel()`; `ECW_automation/main.py::read_patients_from_excel()`.

---

## 19. DOB logic — building reliable age-in-months math on inconsistent Excel date formats

**Symptom/requirement**: the eCW Excel export's DOB and Appointment Date columns weren't perfectly consistent in format across all rows, and simple day-count division (`(appointment_date - dob).days / 30`) is not a reliable way to compute "age in whole months" (calendar months vary in length, and boundary cases near a birthday need to round consistently).

**Final fix**: `utils/date_utils.py::parse_date_flexible()` accepts multiple plausible input shapes (string variants, native `date`/`datetime` objects) and normalizes them to a `date`; `age_in_months()` computes a calendar-aware whole-month difference (not a day-count approximation); `match_asq_bracket()` then applies explicit, deliberately overlap-free cascading ranges (8–10→9, 11–13→12, 14–16→15, 17–20→18, 21–27→24, 28–33→30, 34–41→36, 42–53→48) so every age in the supported range maps to exactly one bracket, with no ambiguous overlap between adjacent brackets.

**Code**: `lone_star_automation/utils/date_utils.py`. Nurture Kids' equivalent logic (`_parse_date_flexible`/`_age_in_months`, locally duplicated inside `ECW_automation/main.py` since that project has no shared `utils` module) follows the same approach.

---

## 20. Cron scheduling / production deployment — honestly, not yet built

**Symptom/requirement**: the documentation request explicitly asked for cron-job details, `daily_run_parallel.sh`, a `PAUSED` file mechanism, Xvfb, virtualenv wrapper, and recovery/error-handling infrastructure — all standard pieces of a real unattended production deployment.

**What's actually true today**: **none of that infrastructure exists in this codebase.** There is no cron job, no shell wrapper script, no `PAUSED` sentinel file, no Xvfb virtual display setup, and no dedicated production virtualenv verified in use. The only two things that exist are `run_production_parallel.py` and `run_demo_parallel.py` — plain Python scripts, launched manually, that run each clinic's own `main.py`/`main_1.py`/`main_demo.py` as a subprocess.

**Why this is documented as a "bug/gap" rather than silently invented**: fabricating a description of infrastructure that doesn't exist would actively mislead a new maintainer taking over this project — they'd go looking for a `daily_run_parallel.sh` or a cron entry that was never created. This is the one part of the requested 12-part documentation structure that's presented as "not yet built," per Part 9 of `PROJECT_DOCUMENTATION.md`, rather than as a resolved bug with a fix.

**Reasonable next step, if built later**: Windows Task Scheduler (since this environment is Windows, not Linux — `cron` doesn't directly apply here) invoking `python main.py` on a recurring schedule; a decision on `headless=True` vs. a Windows-compatible virtual-display equivalent (headless behavior is currently unverified, since several of the resilience fixes above were specifically discovered by watching the visible browser); a simple pause mechanism; and per-project virtual environments for isolation. See Part 9 for the full discussion.

---

## Summary of debugging methodology used throughout

A consistent pattern runs through nearly every entry above, worth naming explicitly for a new maintainer:

1. **Don't guess from the error message alone** — nearly every real fix in this history came from a live, read-only diagnostic script that logged into the real site and dumped actual DOM state (counts, visibility, bounding boxes, screenshots) at the exact point of failure, rather than theorizing from the traceback text.
2. **Distinguish a real code bug from third-party UI flakiness** — a genuine code bug fails the same way, at the same point, every time; flaky third-party UI fails differently each run. Several entries above (items 2, 5, 6, 8) were specifically flakiness, not code defects, and were treated accordingly (retry/stability-poll patterns rather than one-shot "fixes").
3. **When a fix doesn't fully work, say so and show new evidence** — several bugs here (5→6, 14) were solved in two visible passes, where the first pass was honestly reported as insufficient once new live evidence contradicted it, rather than being quietly patched over.
4. **Respect user corrections over your own prior conclusions, especially with screenshots or fresh data** — item 14 is the clearest example: a live diagnostic conclusion was reversed the moment the user supplied contradicting visual evidence, with the mistake and its exact mechanism explained plainly.
