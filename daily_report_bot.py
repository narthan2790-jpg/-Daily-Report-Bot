
import time
import re
import os
from datetime import datetime

import pyautogui
import pyperclip
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# SAFETY: Give yourself time to stop the automation
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

def collect_weather():
    print("Opening Google Chrome...")

    # Open Chrome using Windows Run
    pyautogui.hotkey("win", "r")
    time.sleep(2)
    pyautogui.write("chrome", interval=0.1)
    pyautogui.press("enter")
    time.sleep(5)

    # Visit Chennai weather page
    pyautogui.hotkey("ctrl", "l")
    pyautogui.write("https://wttr.in/Chennai?format=3", interval=0.04)
    pyautogui.press("enter")
    time.sleep(8)

    # Copy webpage text
    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("ctrl", "c")
    time.sleep(2)

    weather_text = pyperclip.paste().strip()

    if not weather_text:
        raise ValueError("Could not copy weather information.")

    print("Weather information collected:")
    print(weather_text)

    # Extract temperature
    match = re.search(r"[+-]?\d+\s*°[CF]", weather_text)

    temperature = match.group(0) if match else "Not found"

    return weather_text, temperature


def create_excel_report(weather_text, temperature):
    today = datetime.now()
    filename = f"Daily_Report_{today.strftime('%Y-%m-%d')}.xlsx"

    # Create workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Daily Status Report"

    # Add heading
    sheet.merge_cells("A1:E1")
    sheet["A1"] = "DAILY STATUS REPORT"
    sheet["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    sheet["A1"].alignment = Alignment(horizontal="center")

    # Add column headings
    headings = [
        "Date & Time",
        "City",
        "Weather Information",
        "Temperature",
        "Daily Comment"
    ]

    for column, heading in enumerate(headings, start=1):
        cell = sheet.cell(row=3, column=column, value=heading)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    # Add report data
    row = [
        today.strftime("%d-%m-%Y %H:%M"),
        "Chennai",
        weather_text,
        temperature,
        "Daily weather status collected automatically using Python."
    ]

    sheet.append([])
    for column, value in enumerate(row, start=1):
        sheet.cell(row=4, column=column, value=value)

    # Adjust column widths
    widths = [22, 15, 45, 18, 65]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[
            chr(64 + index)
        ].width = width

    sheet.freeze_panes = "A4"
    sheet.row_dimensions[4].height = 45

    # Save Excel workbook
    workbook.save(filename)

    print(f"\nExcel report saved: {filename}")

    return os.path.abspath(filename)


def main():
    print("================================")
    print(" DAILY STATUS REPORT AUTOMATION")
    print("================================")

    try:
        weather_text, temperature = collect_weather()

        print("\nCreating Excel report...")
        excel_path = create_excel_report(
            weather_text, temperature
        )

        print("\nOpening Excel report...")
        os.startfile(excel_path)
        time.sleep(8)

        # Save screenshot of the completed Excel screen
        screenshot_path = os.path.abspath(
            "daily_report_screenshot.png"
        )
        pyautogui.screenshot().save(screenshot_path)

        print("\nAutomation completed successfully!")
        print("Excel file:", excel_path)
        print("Screenshot:", screenshot_path)

    except Exception as error:
        print("\nAutomation error:", error)
        print("Please check your internet connection and try again.")


if __name__ == "__main__":
    main()
