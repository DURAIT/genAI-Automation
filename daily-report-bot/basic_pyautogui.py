import os
import time
import datetime as dt
import pyautogui
import pyperclip


# -----------------------------
# Configuration
# -----------------------------
CITY = "Salem"
OUTPUT_DIR = os.path.abspath("daily_reports")

os.makedirs(OUTPUT_DIR, exist_ok=True)

now = dt.datetime.now()
date_str = now.strftime("%Y-%m-%d")
time_str = now.strftime("%Y-%m-%d %H:%M:%S")

excel_file = os.path.join(
    OUTPUT_DIR,
    f"daily_report_{date_str}.xlsx"
)

screenshot_file = os.path.join(
    OUTPUT_DIR,
    f"daily_report_{date_str}.png"
)


# -----------------------------
# Helper functions
# -----------------------------
def wait(seconds=2):
    time.sleep(seconds)


def open_chrome():
    # Windows: launch Chrome through the Start menu.
    pyautogui.hotkey("win", "s")
    wait(1)
    pyautogui.write("Google Chrome", interval=0.03)
    pyautogui.press("enter")
    wait(4)


def get_weather():
    """
    Open Google and search for the weather.
    The temperature is copied from the browser using keyboard navigation.
    """

    pyautogui.hotkey("ctrl", "l")
    pyautogui.write(
        f"https://www.google.com/search?q=weather+{CITY}",
        interval=0.01
    )
    pyautogui.press("enter")
    wait(5)

    # Use browser find to locate the temperature-related text.
    pyautogui.hotkey("ctrl", "f")
    pyautogui.write("°", interval=0.1)
    wait(1)

    pyautogui.press("esc")

    # Select the current page text and copy it.
    # This avoids depending on a fixed screen coordinate.
    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("ctrl", "c")
    wait(1)

    page_text = pyperclip.paste()

    # Look for a temperature such as "32°C" or "90°F".
    import re

    match = re.search(
        r"\b(\d{1,3})\s*°\s*([CF])\b",
        page_text
    )

    if match:
        temperature = f"{match.group(1)}°{match.group(2)}"
    else:
        # Fallback if Google's page layout changes.
        temperature = "Temperature not detected"

    return temperature


def open_excel():
    pyautogui.hotkey("win", "s")
    wait(1)
    pyautogui.write("Microsoft Excel", interval=0.03)
    pyautogui.press("enter")
    wait(5)


def create_excel_report(weather):
    """
    Creates a workbook with:
    Timestamp | Fetched Data | Comment
    """

    # Start a blank workbook.
    pyautogui.hotkey("ctrl", "n")
    wait(3)

    # Enter column headings and the new row.
    data = (
        "Date & Time\tFetched Data\tComment\n"
        f"{time_str}\t{weather}\tGood for outdoor activities"
    )

    pyperclip.copy(data)
    pyautogui.hotkey("ctrl", "v")
    wait(2)

    # Save as XLSX.
    pyautogui.hotkey("ctrl", "shift", "s")
    wait(3)

    # Excel's Save As dialog.
    pyautogui.hotkey("ctrl", "a")
    pyperclip.copy(excel_file)
    pyautogui.hotkey("ctrl", "v")
    wait(1)

    pyautogui.press("enter")
    wait(4)

    # If Excel asks about file format, accept the XLSX format.
    pyautogui.press("enter")
    wait(3)


def take_screenshot():
    """
    Captures the Excel window.
    """

    # Make sure Excel is active.
    pyautogui.hotkey("alt", "tab")
    wait(2)

    # Windows screenshot of the active window.
    # pyautogui.screenshot() captures the whole screen.
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_file)


# -----------------------------
# Main automation
# -----------------------------
def main():

    print("Starting daily report automation...")

    # 1. Open Chrome.
    open_chrome()

    # 2. Fetch important information.
    weather = get_weather()
    print("Fetched data:", weather)

    # 3. Open Excel.
    open_excel()

    # 4. Create and save report.
    create_excel_report(weather)

    # 5. Screenshot final Excel sheet.
    take_screenshot()

    print("\nCompleted.")
    print("Excel file:", excel_file)
    print("Screenshot:", screenshot_file)


if __name__ == "__main__":
    main()

