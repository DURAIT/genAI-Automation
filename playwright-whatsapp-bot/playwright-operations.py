import asyncio
import json
import random
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.async_api import (
    async_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================
# CONFIGURATION
# ============================================================

CONTACTS_FILE = "contacts.xlsx"

# Persistent Chrome profile used by this bot.
# WhatsApp Web login/session information is stored here.
PROFILE_DIR = "whatsapp_profile"

WHATSAPP_URL = "https://web.whatsapp.com/"

# Your installed Google Chrome.
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Playwright timeout values
SELECTOR_TIMEOUT = 20_000
PAGE_TIMEOUT = 30_000

# Delay between normal actions.
MIN_DELAY = 2
MAX_DELAY = 5

# Folder for screenshots
SCREENSHOT_DIR = "screenshots"


# ============================================================
# GENERAL HELPERS
# ============================================================

async def random_delay():
    """
    Random delay between actions.
    This is ordinary rate limiting.
    """
    delay = random.uniform(MIN_DELAY, MAX_DELAY)
    print(f"Waiting {delay:.1f} seconds...")
    await asyncio.sleep(delay)


def get_date():
    return datetime.now().strftime("%Y-%m-%d")


def get_timestamp():
    return datetime.now().isoformat(timespec="seconds")


def clean_phone(phone):
    """
    Normalize phone number.

    Examples:
        +91 98765 43210 -> +919876543210
        +91-98765-43210 -> +919876543210
        919876543210    -> 919876543210
    """

    if phone is None:
        return ""

    phone = str(phone).strip()

    if phone.startswith("+"):
        return "+" + re.sub(r"\D", "", phone[1:])

    return re.sub(r"\D", "", phone)


def safe_filename(value):
    """
    Convert a contact name into a safe filename.
    """
    value = str(value or "contact")
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", value).strip("_")


def personalize_message(template, name):
    """
    Replace {name} in the message template.
    """

    if not template:
        return ""

    return str(template).replace("{name}", str(name))


# ============================================================
# EXCEL INPUT
# ============================================================

def read_contacts():
    """
    Read contacts.xlsx.

    Required columns:
        Name
        Phone

    Optional:
        Message
    """

    file_path = Path(CONTACTS_FILE)

    if not file_path.exists():
        raise FileNotFoundError(
            f"Could not find '{CONTACTS_FILE}'. "
            f"Make sure the Excel file is in the same folder "
            f"as this Python script."
        )

    print(f"Reading {CONTACTS_FILE}...")

    workbook = load_workbook(
        file_path,
        read_only=True,
        data_only=True,
    )

    worksheet = workbook.active

    rows = list(
        worksheet.iter_rows(values_only=True)
    )

    if not rows:
        workbook.close()
        raise ValueError(
            f"{CONTACTS_FILE} is empty."
        )

    headers = [
        str(value).strip()
        if value is not None
        else ""
        for value in rows[0]
    ]

    header_map = {
        header.lower(): index
        for index, header in enumerate(headers)
    }

    if "name" not in header_map:
        workbook.close()
        raise ValueError(
            "Excel file must contain a 'Name' column."
        )

    if "phone" not in header_map:
        workbook.close()
        raise ValueError(
            "Excel file must contain a 'Phone' column."
        )

    message_index = header_map.get("message")

    contacts = []

    for row_number, row in enumerate(
        rows[1:],
        start=2,
    ):

        def get_cell(index):
            if index is None:
                return ""

            if index >= len(row):
                return ""

            return row[index]

        name = str(
            get_cell(header_map["name"]) or ""
        ).strip()

        phone = clean_phone(
            get_cell(header_map["phone"])
        )

        message = ""

        if message_index is not None:
            message = str(
                get_cell(message_index) or ""
            ).strip()

        # Ignore completely blank rows.
        if not name and not phone:
            continue

        contact = {
            "row": row_number,
            "name": name,
            "phone": phone,
            "message_template": message,
        }

        if not phone:
            contact["skip"] = True
            contact["skip_reason"] = (
                "Phone number is missing."
            )
        else:
            contact["skip"] = False

        contacts.append(contact)

    workbook.close()

    return contacts


# ============================================================
# WHATSAPP WEB
# ============================================================

async def wait_for_whatsapp_ready(page):
    """
    Wait for WhatsApp Web.

    On first execution the QR code must be scanned manually.
    """

    print()
    print("Opening WhatsApp Web...")
    print(
        "If this is the first run, scan the QR code manually."
    )
    print()

    await page.goto(
        WHATSAPP_URL,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT,
    )

    search_selectors = [
        'div[contenteditable="true"][data-tab="3"]',
        'div[contenteditable="true"][aria-label*="Search"]',
        'div[role="textbox"][aria-label*="Search"]',
        'input[placeholder*="Search"]',
    ]

    # Wait up to approximately 2 minutes for login.
    for attempt in range(120):

        for selector in search_selectors:

            try:
                locator = page.locator(
                    selector
                ).first

                await locator.wait_for(
                    state="visible",
                    timeout=1_000,
                )

                print()
                print("WhatsApp Web is ready.")
                print()

                return

            except PlaywrightTimeoutError:
                pass

        if attempt % 10 == 0:
            print(
                "Waiting for WhatsApp Web login..."
            )

        await page.wait_for_timeout(1_000)

    raise RuntimeError(
        "WhatsApp Web did not become ready. "
        "Please make sure the QR code was scanned."
    )


# ============================================================
# SEARCH
# ============================================================

async def find_search_box(page):
    """
    Find the WhatsApp search box.
    """

    selectors = [
        'div[contenteditable="true"][data-tab="3"]',
        'div[contenteditable="true"][aria-label*="Search"]',
        'div[role="textbox"][aria-label*="Search"]',
        'input[placeholder*="Search"]',
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            ).first

            await locator.wait_for(
                state="visible",
                timeout=SELECTOR_TIMEOUT,
            )

            return locator

        except PlaywrightTimeoutError:
            continue

    raise RuntimeError(
        "Could not find WhatsApp search box."
    )


async def clear_search_box(search_box):
    """
    Reliably clear the WhatsApp search box.
    """

    await search_box.click()

    try:
        await search_box.fill("")
    except Exception:
        await search_box.press("Control+A")
        await search_box.press("Backspace")


async def search_contact(page, name, phone):
    """
    Search WhatsApp by phone number first,
    then by name.

    Returns:
        True  -> contact/chat found
        False -> contact/chat not found
    """

    search_box = await find_search_box(page)

    search_terms = []

    if phone:
        search_terms.append(phone)

        # Also try phone without '+'
        phone_without_plus = phone.lstrip("+")

        if phone_without_plus != phone:
            search_terms.append(
                phone_without_plus
            )

    if name:
        search_terms.append(name)

    for search_term in search_terms:

        try:

            print(
                f"Searching: {search_term}"
            )

            await clear_search_box(
                search_box
            )

            await search_box.type(
                search_term,
                delay=40,
            )

            await random_delay()

            # Allow search results to populate.
            await page.wait_for_timeout(
                1_500
            )

            # ----------------------------------------
            # Try common result containers
            # ----------------------------------------

            result_selectors = [
                '[role="listitem"]',
                'div[role="option"]',
                'div[aria-label*="Chat"]',
            ]

            for selector in result_selectors:

                results = page.locator(
                    selector
                )

                try:
                    count = await results.count()
                except Exception:
                    count = 0

                if count == 0:
                    continue

                for index in range(
                    min(count, 15)
                ):

                    result = results.nth(index)

                    try:

                        text = await result.inner_text(
                            timeout=1_000
                        )

                    except Exception:
                        continue

                    normalized = (
                        text.lower()
                    )

                    term_matches = (
                        str(search_term).lower()
                        in normalized
                    )

                    name_matches = (
                        bool(name)
                        and name.lower()
                        in normalized
                    )

                    if term_matches or name_matches:

                        print(
                            "Contact search result found."
                        )

                        await result.click()

                        await page.wait_for_timeout(
                            1_000
                        )

                        return True

            # ----------------------------------------
            # Fallback: exact name
            # ----------------------------------------

            if name:

                try:

                    name_locator = page.get_by_text(
                        name,
                        exact=True,
                    ).first

                    await name_locator.click(
                        timeout=2_000
                    )

                    await page.wait_for_timeout(
                        1_000
                    )

                    return True

                except Exception:
                    pass

        except Exception as exc:

            print(
                f"Search attempt failed: {exc}"
            )

            continue

    return False


# ============================================================
# MESSAGE BOX
# ============================================================

async def find_message_box(page):
    """
    Locate WhatsApp's 'Type a message' field.
    """

    selectors = [
        'div[contenteditable="true"][aria-label="Type a message"]',
        'div[contenteditable="true"][data-tab="10"]',
        'div[contenteditable="true"][data-tab="6"]',
        'footer div[contenteditable="true"]',
        'div[role="textbox"][contenteditable="true"]',
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            ).last

            await locator.wait_for(
                state="visible",
                timeout=SELECTOR_TIMEOUT,
            )

            return locator

        except PlaywrightTimeoutError:
            continue

    raise RuntimeError(
        "Could not find WhatsApp message box."
    )


async def send_message(page, message):
    """
    Type and send a message.
    """

    if not message.strip():
        raise ValueError(
            "Message is empty."
        )

    message_box = await find_message_box(
        page
    )

    await message_box.click()

    # Try fill first.
    try:

        await message_box.fill(
            message
        )

    except Exception:

        # Fallback for contenteditable.
        await message_box.press(
            "Control+A"
        )

        await message_box.type(
            message,
            delay=20,
        )

    await random_delay()

    # Send.
    await message_box.press(
        "Enter"
    )

    # Give WhatsApp time to render
    # the outgoing message.
    await page.wait_for_timeout(
        2_000
    )

    return True


# ============================================================
# SCREENSHOT
# ============================================================

async def screenshot_sent_message(
    page,
    name,
    index,
):
    """
    Screenshot the current WhatsApp chat
    after sending.
    """

    screenshot_directory = Path(
        SCREENSHOT_DIR
    )

    screenshot_directory.mkdir(
        exist_ok=True
    )

    filename = (
        f"{get_date()}_"
        f"{index:04d}_"
        f"{safe_filename(name)}.png"
    )

    filepath = (
        screenshot_directory
        / filename
    )

    await page.screenshot(
        path=str(filepath),
        full_page=False,
    )

    return str(filepath)


# ============================================================
# LAST 3 MESSAGES
# ============================================================

async def extract_last_three_messages(page):
    """
    Extract the last 3 incoming messages.

    WhatsApp Web DOM selectors can change, so several
    selectors are attempted.
    """

    selectors = [
        "div.message-in",
        'div[data-id][class*="message-in"]',
    ]

    messages = []

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            )

            count = await locator.count()

            if count == 0:
                continue

            # Start from newest.
            for index in range(
                count - 1,
                max(-1, count - 30),
                -1,
            ):

                message = locator.nth(
                    index
                )

                try:

                    text = await message.inner_text(
                        timeout=1_000
                    )

                except Exception:
                    continue

                text = " ".join(
                    text.split()
                ).strip()

                if not text:
                    continue

                if text in messages:
                    continue

                messages.append(text)

                if len(messages) == 3:
                    break

            if messages:
                break

        except Exception:
            continue

    # Currently newest -> oldest.
    # Return oldest -> newest.
    messages.reverse()

    return messages[-3:]


# ============================================================
# REPORTING
# ============================================================

def save_json_report(results):
    """
    Save full JSON report.
    """

    filename = (
        f"whatsapp_report_{get_date()}.json"
    )

    report = {
        "generated_at": get_timestamp(),
        "contacts_file": CONTACTS_FILE,
        "total_contacts": len(results),
        "successful": sum(
            1
            for item in results
            if item.get("status") == "sent"
        ),
        "failed": sum(
            1
            for item in results
            if item.get("status") == "failed"
        ),
        "skipped": sum(
            1
            for item in results
            if item.get("status") == "skipped"
        ),
        "contacts": results,
    }

    with open(
        filename,
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            report,
            file,
            ensure_ascii=False,
            indent=2,
        )

    return filename


def save_excel_report(results):
    """
    Save summary Excel report.
    """

    filename = (
        f"whatsapp_report_{get_date()}.xlsx"
    )

    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = "Summary"

    headers = [
        "Name",
        "Phone",
        "Status",
        "Message",
        "Screenshot",
        "Last 3 Messages",
        "Error",
        "Timestamp",
    ]

    worksheet.append(headers)

    for result in results:

        last_messages = result.get(
            "last_3_messages",
            [],
        )

        worksheet.append([
            result.get("name", ""),
            result.get("phone", ""),
            result.get("status", ""),
            result.get("message", ""),
            result.get("screenshot", ""),
            "\n".join(last_messages),
            result.get("error", ""),
            result.get("timestamp", ""),
        ])

    # Formatting
    for column in worksheet.columns:

        max_length = 0

        column_letter = (
            column[0].column_letter
        )

        for cell in column:

            value = str(
                cell.value or ""
            )

            max_length = max(
                max_length,
                len(value),
            )

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max(max_length + 2, 12),
            60,
        )

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    workbook.save(filename)

    return filename


# ============================================================
# PROCESS ONE CONTACT
# ============================================================

async def process_contact(
    page,
    contact,
    index,
):
    """
    Process one contact.

    Any error is captured in the result rather than
    crashing the entire program.
    """

    name = contact.get(
        "name",
        "",
    )

    phone = contact.get(
        "phone",
        "",
    )

    template = contact.get(
        "message_template",
        "",
    )

    result = {
        "row": contact.get("row"),
        "name": name,
        "phone": phone,
        "message": "",
        "status": "failed",
        "screenshot": "",
        "last_3_messages": [],
        "error": "",
        "timestamp": get_timestamp(),
    }

    # ----------------------------------------
    # Skip invalid contact
    # ----------------------------------------

    if contact.get("skip"):

        result["status"] = "skipped"

        result["error"] = contact.get(
            "skip_reason",
            "Skipped.",
        )

        return result

    # ----------------------------------------
    # Personalize message
    # ----------------------------------------

    message = personalize_message(
        template,
        name,
    )

    if not message:

        result["status"] = "skipped"

        result["error"] = (
            "No message supplied."
        )

        return result

    result["message"] = message

    try:

        print()
        print(
            "=" * 60
        )

        print(
            f"Contact {index}: "
            f"{name} ({phone})"
        )

        print(
            f"Message: {message}"
        )

        print(
            "=" * 60
        )

        # ----------------------------------------
        # Search
        # ----------------------------------------

        found = await search_contact(
            page,
            name,
            phone,
        )

        if not found:

            result["status"] = "failed"

            result["error"] = (
                "Contact not found."
            )

            print(
                f"FAILED: {name} - contact not found."
            )

            return result

        await random_delay()

        # ----------------------------------------
        # Send
        # ----------------------------------------

        print(
            f"Sending message to {name}..."
        )

        await send_message(
            page,
            message,
        )

        # ----------------------------------------
        # Screenshot
        # ----------------------------------------

        await page.wait_for_timeout(
            2_000
        )

        screenshot = (
            await screenshot_sent_message(
                page,
                name,
                index,
            )
        )

        result["screenshot"] = screenshot

        print(
            f"Screenshot saved: {screenshot}"
        )

        # ----------------------------------------
        # Extract last 3 messages
        # ----------------------------------------

        await random_delay()

        print(
            "Extracting last 3 messages..."
        )

        last_messages = (
            await extract_last_three_messages(
                page
            )
        )

        result["last_3_messages"] = (
            last_messages
        )

        # ----------------------------------------
        # Success
        # ----------------------------------------

        result["status"] = "sent"

        result["timestamp"] = (
            get_timestamp()
        )

        print(
            f"SUCCESS: {name}"
        )

    except PlaywrightTimeoutError as exc:

        result["status"] = "failed"

        result["error"] = (
            "Playwright timeout: "
            + str(exc)
        )

        print(
            f"TIMEOUT: {name}"
        )

    except Exception as exc:

        result["status"] = "failed"

        result["error"] = (
            f"{type(exc).__name__}: "
            f"{str(exc)}"
        )

        print(
            f"ERROR processing {name}: "
            f"{exc}"
        )

    return result


# ============================================================
# MAIN
# ============================================================

async def main():

    try:

        # ----------------------------------------
        # Verify Chrome
        # ----------------------------------------

        chrome_path = Path(
            CHROME_PATH
        )

        if not chrome_path.exists():

            raise FileNotFoundError(
                "Google Chrome was not found at:\n"
                f"{CHROME_PATH}\n\n"
                "Please update CHROME_PATH in "
                "this script."
            )

        print(
            f"Using Chrome:\n{CHROME_PATH}"
        )

        # ----------------------------------------
        # Read Excel
        # ----------------------------------------

        contacts = read_contacts()

        if not contacts:

            print(
                "No contacts found."
            )

            return

        print(
            f"Loaded {len(contacts)} contacts."
        )

        results = []

        # ----------------------------------------
        # Start Playwright
        # ----------------------------------------

        async with async_playwright() as p:

            print()
            print(
                "Starting Google Chrome..."
            )

            browser = (
                await p.chromium.launch_persistent_context(
                    user_data_dir=PROFILE_DIR,

                    # IMPORTANT:
                    # Use installed Google Chrome instead
                    # of Playwright's downloaded Chromium.
                    executable_path=CHROME_PATH,

                    headless=False,

                    viewport={
                        "width": 1440,
                        "height": 1000,
                    },

                    args=[
                        "--start-maximized",
                    ],
                )
            )

            # Reuse existing page if available.
            if browser.pages:

                page = browser.pages[0]

            else:

                page = await browser.new_page()

            page.set_default_timeout(
                SELECTOR_TIMEOUT
            )

            page.set_default_navigation_timeout(
                PAGE_TIMEOUT
            )

            try:

                # ----------------------------------------
                # WhatsApp login
                # ----------------------------------------

                await wait_for_whatsapp_ready(
                    page
                )

                # ----------------------------------------
                # Process contacts
                # ----------------------------------------

                for index, contact in enumerate(
                    contacts,
                    start=1,
                ):

                    result = (
                        await process_contact(
                            page,
                            contact,
                            index,
                        )
                    )

                    results.append(result)

                    # ------------------------------------
                    # Save reports after every contact.
                    # This protects earlier results if
                    # something unexpected happens later.
                    # ------------------------------------

                    json_file = (
                        save_json_report(
                            results
                        )
                    )

                    excel_file = (
                        save_excel_report(
                            results
                        )
                    )

                    print(
                        f"Progress saved:"
                    )

                    print(
                        f"  {json_file}"
                    )

                    print(
                        f"  {excel_file}"
                    )

                    # Delay before next contact.
                    if index < len(contacts):

                        await random_delay()

            finally:

                # ----------------------------------------
                # Always save final reports
                # ----------------------------------------

                print()
                print(
                    "Saving final reports..."
                )

                json_file = (
                    save_json_report(
                        results
                    )
                )

                excel_file = (
                    save_excel_report(
                        results
                    )
                )

                print()
                print(
                    "=" * 60
                )

                print(
                    "FINAL REPORTS"
                )

                print(
                    "=" * 60
                )

                print(
                    f"JSON : {json_file}"
                )

                print(
                    f"Excel: {excel_file}"
                )

                print(
                    "=" * 60
                )

                await browser.close()

    except Exception as exc:

        print()
        print(
            "FATAL ERROR:"
        )

        print(
            f"{type(exc).__name__}: {exc}"
        )

        raise


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
