"""
ecw/facility_filter.py
------------------------
Filters the Encounter Patient Download report to "Lone Star Pediatrics
Midlothian" via the report screen's Facility tab, before the date range
is set.

Selectors below were reverse-engineered from a live capture of the real
report screen (logs/inspect/report_iframe.html /report_screen.png via
ecw/inspect_facility_screen.py), not guessed blindly:

  - Tabs (#sns-tabs) are plain <span> elements, not ARIA role="tab" -
    "Facility" is already the default-active tab, but we click it anyway
    to match the confirmed workflow (harmless no-op if already active).
  - Each filter section (Facility Name, Facility POS, etc.) is a
    collapsed <div class="sns-container" style="display:none"> toggled
    open by clicking a sibling <div id="facility-label"> - that id is
    static, unlike the randomized per-widget instance ids inside it
    (e.g. "N0x117cf380x0x1167c368_NS_"), so we scope all inner locators
    off the static "#facility-container" instead of those ids.
  - The Results/Choice lists are native <select multiple size="9">
    listboxes (MicroStrategy's "selectWithSearch" prompt widget) - they
    render inline (not as an OS popup), so options are directly
    clickable.
  - IMPORTANT: there is exactly ONE "OK"/Finish button for the entire
    report screen (facility + dates together) - it lives in the page
    footer, not inside this filter. Do NOT click it here, or the report
    would submit before the date range is set. schedule_export.py clicks
    it once, after both facility and dates are configured, unchanged
    from the reference project's behavior.
"""

import asyncio

from config import settings
from utils.logger import get_logger

log = get_logger(__name__)


async def apply_facility_filter(iframe):
    log.info("Applying Facility filter...")

    # Facility is the default-active tab already, but click it explicitly
    # to match the confirmed workflow - harmless if already selected.
    # force=True: an empty sibling <span class="counts"> badge visually
    # overlaps the tab label and blocks Playwright's normal actionability
    # check even though it has no click handler of its own.
    await iframe.locator("#sns-tabs").get_by_text("Facility", exact=True).click(force=True)
    await asyncio.sleep(0.5)
    log.info("Facility tab active.")

    # Expand the "Facility Name" filter by clicking the "+" cell specifically
    # (the onclick="EBO.common.toggle(...)" handler lives on the <td> cells,
    # not on the wrapping #facility-label div itself - clicking the div's
    # bounding-box center isn't guaranteed to land on either cell)
    await iframe.locator("#facility-label td").first.click()
    log.info("Facility Name filter opened (+ clicked).")

    container = iframe.locator("#facility-container")
    await container.wait_for(state="visible", timeout=10000)

    keywords_input = container.locator('input[name="_searchValue"]')
    await keywords_input.fill(settings.FACILITY_KEYWORD)

    search_button = container.get_by_role("button", name="Search")
    await search_button.click()
    log.info(f"Searched Facility keyword: '{settings.FACILITY_KEYWORD}'")

    results_select = container.locator('select[id^="PRMT_SV_"]')
    await results_select.wait_for(state="visible", timeout=10000)
    options = results_select.locator("option")

    # The Results list populates asynchronously after Search - a fixed sleep
    # raced ahead of it once (only "Lone Star Pediatrics" had rendered, not
    # yet "...Midlothian", so the wrong one got inserted). Poll until the
    # option count stops changing instead of guessing a fixed delay.
    previous_count = -1
    stable_checks = 0
    count = 0
    for _ in range(20):
        count = await options.count()
        if count == previous_count and count > 0:
            stable_checks += 1
            if stable_checks >= 2:
                break
        else:
            stable_checks = 0
        previous_count = count
        await asyncio.sleep(0.5)

    option_texts = await options.all_inner_texts()
    log.info(f"Found {count} facility result(s): {option_texts}")

    # Always select the SECOND result (Lone Star Pediatrics Midlothian)
    if count >= 2:
        target_index = 1
    else:
        log.info("Expected 2 results, found fewer - selecting the only available match.")
        target_index = 0

    await options.nth(target_index).click()
    log.info(f"Selected result: '{option_texts[target_index] if option_texts else '?'}'")

    insert_button = container.get_by_role("button", name="Insert")
    await insert_button.click()
    await asyncio.sleep(1)
    log.info("Facility inserted into Choice list.")

    # Verify it landed in the Choice list before moving on to dates
    choice_select = container.locator('select[id^="PRMT_LIST_BOX_SELECT_"]')
    choice_texts = await choice_select.locator("option").all_inner_texts()
    log.info(f"Choice list now contains: {choice_texts}")
    if not any(settings.FACILITY_NAME in t for t in choice_texts):
        raise RuntimeError(
            f"'{settings.FACILITY_NAME}' not found in Choice list after Insert: {choice_texts}"
        )
    log.info(f"Verified '{settings.FACILITY_NAME}' present in Choice list.")
    # NOTE: no OK click here - see module docstring.
