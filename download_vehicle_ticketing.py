"""Download monthly Vehicle Ticketing exports from Whizzard EV.

Install:
    pip install playwright
    playwright install chromium

Run:
    python download_vehicle_ticketing.py

This automation is read-only: it never raises, edits, clears, refreshes, or closes a ticket.
"""

from __future__ import annotations

import csv
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from playwright.sync_api import (
    Browser,
    Download,
    Locator,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)


# ------------------------------- CONFIG -----------------------------------
# Keep credentials here as requested or read from environment variables.
WHIZZARD_USER = (os.getenv("WHIZZARD_USER") or "").strip() or "7658978951"
WHIZZARD_PASS = (os.getenv("WHIZZARD_PASS") or "").strip() or "7658978951"

BASE_URL = "https://whizzard-ev.com"
OUTPUT_PATH = Path("./Whizzard_Downloads")
HEADLESS = False
OVERWRITE_EXISTING = False

# Accepted examples: "Sep 2025", "September 2025", "2025-09",
#                    "Sep 2025 to Oct 2026", or "ALL".
MONTH_SELECTION = "Sep 2025 to Oct 2026"

VALID_START = (2025, 9)
VALID_END = (2026, 10)
DOWNLOAD_TIMEOUT_MS = 15_000
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Section:
    tab: str
    file_stem: str
    folder: Path
    action: str
    subtab: str | None = None


SECTIONS = (
    Section("Return from Ops", "Return_from_Ops", Path("1_Return_from_Ops"), "Search"),
    Section("Handover to Ops", "Handover_to_Ops", Path("2_Handover_to_Ops"), "Search"),
    Section("Service Request", "Service_Request", Path("3_Service_Request"), "Search"),
    Section(
        "Vehicle Movement", "Inter_Station_Transfer",
        Path("4_Vehicle_Movement/Inter_Station_Transfer"), "Apply filters",
        "Inter Station transfer",
    ),
    Section(
        "Vehicle Movement", "Vehicle_Requisition",
        Path("4_Vehicle_Movement/Vehicle_Requisition"), "Apply filters",
        "Vehicle Requisition",
    ),
)

MONTH_ALIASES = {name.lower(): i for i, name in enumerate(
    ("January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"), 1
)}
MONTH_ALIASES.update({name[:3].lower(): number for name, number in MONTH_ALIASES.copy().items()})


def parse_month(text: str) -> tuple[int, int]:
    value = text.strip()
    iso = re.fullmatch(r"(\d{4})-(\d{1,2})", value)
    named = re.fullmatch(r"([A-Za-z]+)\s+(\d{4})", value)
    if iso:
        year, month = int(iso.group(1)), int(iso.group(2))
    elif named and named.group(1).lower() in MONTH_ALIASES:
        month, year = MONTH_ALIASES[named.group(1).lower()], int(named.group(2))
    else:
        raise ValueError(f"Unrecognized month: {text!r}")
    if not 1 <= month <= 12:
        raise ValueError(f"Invalid month: {text!r}")
    return year, month


def month_number(value: tuple[int, int]) -> int:
    return value[0] * 12 + value[1]


def iter_months(start: tuple[int, int], end: tuple[int, int]) -> Iterable[tuple[int, int]]:
    year, month = start
    while (year, month) <= end:
        yield year, month
        month += 1
        if month == 13:
            year, month = year + 1, 1


def build_months(selection: str) -> list[tuple[int, int]]:
    value = selection.strip()
    if value.lower() in {"all", "3"}:
        start, end = VALID_START, VALID_END
    else:
        parts = re.split(r"\s+to\s+", value, maxsplit=1, flags=re.IGNORECASE)
        start = parse_month(parts[0])
        end = parse_month(parts[1]) if len(parts) == 2 else start
    if start > end:
        raise ValueError("The start month must not be after the end month.")
    if month_number(start) < month_number(VALID_START) or month_number(end) > month_number(VALID_END):
        raise ValueError("Months must be between September 2025 and October 2026.")
    return list(iter_months(start, end))


def month_label(month: tuple[int, int]) -> str:
    return datetime(month[0], month[1], 1).strftime("%B %Y")


def month_iso(month: tuple[int, int]) -> str:
    return f"{month[0]:04d}-{month[1]:02d}"


def first_visible(*locators: Locator) -> Locator:
    for locator in locators:
        try:
            for index in range(locator.count()):
                candidate = locator.nth(index)
                if candidate.is_visible():
                    return candidate
        except Exception:
            pass
    raise RuntimeError("No matching visible element was found.")


def screenshot_failure(page: Page, debug_dir: Path, name: str) -> None:
    """Capture state without ever preserving a password value."""
    try:
        password = page.locator('input[type="password"]')
        for index in range(password.count()):
            if password.nth(index).is_visible():
                password.nth(index).fill("")
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", name)
        page.screenshot(path=str(debug_dir / f"{safe}.png"), full_page=True)
    except Exception:
        pass


def login(page: Page, debug_dir: Path) -> None:
    page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60_000)
    mobile = first_visible(
        page.get_by_label("Mobile Number", exact=False),
        page.locator('input[placeholder*="mobile" i]'),
        page.locator('input[type="tel"]'),
    )
    password = first_visible(
        page.get_by_label("Password", exact=False),
        page.locator('input[type="password"]'),
    )
    mobile.fill(WHIZZARD_USER)
    password.fill(WHIZZARD_PASS)
    page.get_by_role("button", name=re.compile(r"Sign In", re.I)).click()
    try:
        page.get_by_text("Fleet At a Glance", exact=False).wait_for(state="visible", timeout=30_000)
    except Exception as exc:
        # Clear the secret before the required debug screenshot.
        screenshot_failure(page, debug_dir, "login_failed")
        raise RuntimeError("Login failed: 'Fleet At a Glance' did not appear.") from exc


def ensure_logged_in(page: Page, debug_dir: Path) -> None:
    login_controls = page.locator('input[type="password"]')
    if "/login" in page.url.lower() or (login_controls.count() and login_controls.first.is_visible()):
        login(page, debug_dir)


def open_module(page: Page) -> None:
    page.get_by_text("Vehicle Ticketing", exact=True).first.click()
    page.get_by_text("Vehicle Ticketing System", exact=False).wait_for(state="visible", timeout=20_000)


def wait_for_loading(page: Page) -> None:
    """Wait for the module's blocking loading overlay to disappear."""
    try:
        page.wait_for_function(
            """() => !/Loading (tickets|handover tickets|service requests|transfers|requisitions)/i
                .test(document.body.innerText)""",
            timeout=30_000,
        )
    except PlaywrightTimeoutError:
        pass


def click_section(page: Page, section: Section) -> None:
    first_visible(page.get_by_text(section.tab, exact=True)).click()
    if section.subtab:
        first_visible(page.get_by_text(section.subtab, exact=True)).click()
    first_visible(page.locator("table")).wait_for(state="visible", timeout=20_000)
    wait_for_loading(page)


def find_month_field(page: Page, section: Section) -> Locator:
    label = "Month" if section.subtab else "Filter by Month"
    candidates = (
        page.get_by_label(label, exact=True),
        page.locator(f'label:has-text("{label}")').locator("input"),
        page.locator('input[type="month"]'),
        page.locator('input[placeholder*="month" i]'),
    )
    return first_visible(*candidates)


def set_month(page: Page, field: Locator, target: tuple[int, int]) -> None:
    """Handle native month inputs, flatpickr, or text-backed custom date fields."""
    iso, label = month_iso(target), month_label(target)
    input_type = (field.get_attribute("type") or "").lower()
    if input_type == "month":
        field.fill(iso)
    elif field.get_attribute("data-lw-month") is not None or field.get_attribute("readonly") is not None:
        field.click()
        picker = page.locator(".air-datepicker:visible")
        picker.wait_for(state="visible", timeout=5_000)
        title = picker.locator(".air-datepicker-nav--title")
        for _ in range(20):
            shown_year = int(re.search(r"\d{4}", title.inner_text()).group())
            if shown_year == target[0]:
                break
            direction = "prev" if shown_year > target[0] else "next"
            picker.locator(f'.air-datepicker-nav--action[data-action="{direction}"]').click()
        else:
            raise RuntimeError(f"Could not navigate month picker to {target[0]}.")
        picker.locator(
            f'.air-datepicker-cell.-month-[data-year="{target[0]}"][data-month="{target[1] - 1}"]'
        ).click()
    else:
        # Flatpickr instances need their own setter so framework state is updated.
        used_flatpickr = field.evaluate(
            """(el, value) => {
                if (el._flatpickr) { el._flatpickr.setDate(value, true, 'Y-m'); return true; }
                return false;
            }""",
            iso,
        )
        if not used_flatpickr:
            try:
                field.fill(label)
            except Exception:
                # React/custom readonly inputs: set the native value and emit both events.
                field.evaluate(
                    """(el, value) => {
                        const setter = Object.getOwnPropertyDescriptor(
                            HTMLInputElement.prototype, 'value').set;
                        setter.call(el, value);
                        el.dispatchEvent(new Event('input', {bubbles: true}));
                        el.dispatchEvent(new Event('change', {bubbles: true}));
                    }""",
                    label,
                )
    page.wait_for_timeout(1_000)
    shown = (field.input_value() or "").strip()
    if iso not in shown and label.lower() not in shown.lower():
        raise RuntimeError(f"Month filter verification failed; field shows {shown!r}.")


def table_signature(page: Page) -> tuple[int, str]:
    rows = page.locator("table:visible tbody tr")
    count = rows.count()
    text = rows.all_inner_texts()
    return count, "|".join(text)[:2000]


def apply_filter(page: Page, section: Section) -> int:
    before = table_signature(page)
    action = first_visible(
        page.get_by_text(section.action, exact=True),
        page.get_by_role("button", name=re.compile(f"^{re.escape(section.action)}$", re.I)),
    )
    action.click()
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PlaywrightTimeoutError:
        pass
    wait_for_loading(page)
    try:
        page.wait_for_function(
            """([count, text]) => {
                const rows = [...document.querySelectorAll('table tbody tr')];
                return rows.length !== count || rows.map(r => r.innerText).join('|').slice(0,2000) !== text;
            }""",
            arg=list(before), timeout=8_000,
        )
    except PlaywrightTimeoutError:
        pass
    rows = page.locator("table:visible tbody tr")
    empty = re.compile(r"No\s+(tickets?|service request tickets?|inter station transfers?|vehicle requisitions?)", re.I)
    return sum(1 for text in rows.all_inner_texts() if text.strip() and not empty.search(text))


def versioned_path(path: Path) -> Path:
    if not path.exists():
        return path
    number = 2
    while True:
        candidate = path.with_name(f"{path.stem}_v{number}{path.suffix}")
        if not candidate.exists():
            return candidate
        number += 1


def close_status_modal(page: Page) -> None:
    """Dismiss only the export status dialog; this never changes ticket state."""
    modal = page.locator("#statusModal.active")
    if modal.count() and modal.is_visible():
        first_visible(modal.get_by_role("button", name="Close", exact=True)).click()
        modal.wait_for(state="hidden", timeout=5_000)


def save_download(download: Download, section: Section, target: tuple[int, int]) -> Path:
    suggested_ext = Path(download.suggested_filename).suffix
    extension = suggested_ext if suggested_ext else ".xlsx"
    destination = OUTPUT_PATH / section.folder / f"{section.file_stem}_{month_iso(target)}{extension}"
    if not OVERWRITE_EXISTING:
        destination = versioned_path(destination)
    download.save_as(str(destination.resolve()))
    return destination


def export_once(page: Page, section: Section, target: tuple[int, int]) -> Path:
    export = first_visible(
        page.get_by_text("Export", exact=True),
        page.get_by_role("button", name=re.compile(r"^Export$", re.I)),
    )
    with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as event:
        export.click()
    saved = save_download(event.value, section, target)
    close_status_modal(page)
    return saved


def process_one(page: Page, debug_dir: Path, section: Section, target: tuple[int, int]) -> dict[str, str | int]:
    label = month_label(target)
    row_count = 0
    for attempt in range(1, 4):
        try:
            ensure_logged_in(page, debug_dir)
            try:
                first_visible(page.get_by_text("Vehicle Ticketing System", exact=False))
            except RuntimeError:
                open_module(page)
            click_section(page, section)
            field = find_month_field(page, section)
            set_month(page, field, target)
            row_count = apply_filter(page, section)
            saved = export_once(page, section, target)
            size = saved.stat().st_size
            return {"Section": section.file_stem, "Month": label, "File saved": str(saved),
                    "Rows": row_count, "Size bytes": size, "Status": "OK"}
        except PlaywrightTimeoutError as exc:
            print(f"Attempt {attempt} timed out for {section.file_stem} {label}: {exc}", flush=True)
            page_text = page.locator("body").inner_text(timeout=5_000)
            no_data = bool(re.search(
                r"No (tickets found|service request tickets found|inter station transfers yet|vehicle requisitions yet)",
                page_text, re.I,
            ))
            if no_data:
                close_status_modal(page)
                return {"Section": section.file_stem, "Month": label, "File saved": "",
                        "Rows": 0, "Size bytes": 0, "Status": "NO DATA"}
            screenshot_failure(page, debug_dir, f"{section.file_stem}_{month_iso(target)}_attempt{attempt}")
            close_status_modal(page)
        except Exception as exc:
            print(
                f"Attempt {attempt} failed for {section.file_stem} {label}: "
                f"{type(exc).__name__}: {exc}", flush=True,
            )
            screenshot_failure(page, debug_dir, f"{section.file_stem}_{month_iso(target)}_attempt{attempt}")
            close_status_modal(page)
        if attempt < 3:
            try:
                ensure_logged_in(page, debug_dir)
            except Exception:
                pass
    return {"Section": section.file_stem, "Month": label, "File saved": "",
            "Rows": row_count, "Size bytes": 0, "Status": "FAILED"}


def prepare_folders() -> Path:
    for section in SECTIONS:
        (OUTPUT_PATH / section.folder).mkdir(parents=True, exist_ok=True)
    debug = OUTPUT_PATH / "_debug"
    debug.mkdir(parents=True, exist_ok=True)
    return debug


def write_log(results: list[dict[str, str | int]]) -> None:
    log_path = OUTPUT_PATH / "download_log.csv"
    fields = ["Section", "Month", "File saved", "Rows", "Size bytes", "Status"]
    with log_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(results)


def print_summary(results: list[dict[str, str | int]]) -> None:
    print("\nSection | Month | File saved | Status")
    print("-" * 100)
    for item in results:
        print(f"{item['Section']} | {item['Month']} | {item['File saved']} | {item['Status']}")
    ok = sum(item["Status"] == "OK" for item in results)
    skipped = sum(item["Status"] == "NO DATA" for item in results)
    failed = sum(item["Status"] == "FAILED" for item in results)
    print(f"\nDownloaded: {ok} | Skipped (no data): {skipped} | Failed: {failed}")


def main() -> int:
    try:
        months = build_months(MONTH_SELECTION)
    except ValueError as exc:
        print(f"Invalid MONTH_SELECTION: {exc}", file=sys.stderr)
        return 2

    debug_dir = prepare_folders()
    results: list[dict[str, str | int]] = []
    with sync_playwright() as playwright:
        browser: Browser = playwright.chromium.launch(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1440, "height": 900},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        )
        page = context.new_page()
        try:
            login(page, debug_dir)
            open_module(page)
            for section in SECTIONS:
                for target in months:
                    result = process_one(page, debug_dir, section, target)
                    results.append(result)
                    print(f"{result['Section']} | {result['Month']} | {result['Status']}")
                    write_log(results)  # Preserve progress after every attempt.
        finally:
            browser.close()

    write_log(results)
    print_summary(results)
    return 0 if all(item["Status"] != "FAILED" for item in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
