import os
import pyautogui
import time
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from datetime import datetime

# Step 1: Open the Excel file in its default application
def open_excel_file(file_path):
    if os.name == 'nt':  # Windows
        os.startfile(file_path)
    elif os.name == 'posix':  # macOS/Linux
        os.system(f'open "{file_path}"')

# Step 2: Simulate column selection in Excel
def select_column_in_excel(column_letter="A"):
    time.sleep(5)  # Wait for Excel to fully open
    # Navigate to cell A2
    pyautogui.hotkey("control", "g")  # Open 'Go To' dialog (Command+G on macOS)
    pyautogui.typewrite("A2")  # Specify cell A2
    pyautogui.press("return")  # Press Return (Enter) to go to A2

    time.sleep(1)  # Short delay to ensure A2 is selected
    # Select all rows from A2 downwards
    pyautogui.keyDown("command")  # Hold Command
    pyautogui.keyDown("shift")  # Hold Shift
    pyautogui.press("down")  # Press Down Arrow
    pyautogui.keyUp("shift")  # Release Shift
    pyautogui.keyUp("command")  # Release Command
    time.sleep(5)

# Step 3: Take a screenshot of the Excel window
def take_screenshot(save_path):
    screenshot = pyautogui.screenshot()
    screenshot.save(save_path)

# Step 4: Attach screenshots to a single Excel file
def attach_screenshots_to_excel(screenshot_paths, output_excel_path):
    # Create a new workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Screenshots"

    # Insert each screenshot into a new row
    row = 2  # Start placing images from row 2
    for idx, screenshot_path in enumerate(screenshot_paths, start=1):
        # Add a label for the screenshot
        sheet.cell(row=row, column=1, value=f"Screenshot {idx}")

        # Insert the screenshot image
        img = Image(screenshot_path)
        img.anchor = f"B{row}"  # Place the image in column B
        sheet.add_image(img)

        # Move to the next row for the next screenshot
        row += 20  # Adjust spacing for images

    # Save the final Excel file
    workbook.save(output_excel_path)
    print(f"All screenshots saved in: {output_excel_path}")

# Main function to process multiple Excel files
def automate_multiple_excel_screenshots(excel_file_paths, column_letter="A"):
    # Prepare output paths
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel_path = f"aggregated_screenshots_{timestamp}.xlsx"
    screenshot_folder = "screenshots"
    os.makedirs(screenshot_folder, exist_ok=True)

    screenshot_paths = []  # To store paths of all screenshots

    for file_path in excel_file_paths:
        print(f"Processing file: {file_path}")
        # Open the Excel file
        open_excel_file(file_path)

        # Wait for Excel to open and select the column
        print("Waiting for Excel to open...")
        time.sleep(10)
        print(f"Selecting column {column_letter} in Excel...")
        select_column_in_excel(column_letter)

        # Take a screenshot
        screenshot_name = f"{os.path.splitext(os.path.basename(file_path))[0]}_{timestamp}.png"
        screenshot_path = os.path.join(screenshot_folder, screenshot_name)
        print(f"Taking a screenshot for: {file_path}")
        take_screenshot(screenshot_path)
        screenshot_paths.append(screenshot_path)

        print(f"Screenshot saved: {screenshot_path}")

    # Save all screenshots into a single Excel file
    print("Saving all screenshots to a single Excel file...")
    attach_screenshots_to_excel(screenshot_paths, output_excel_path)

    print("Process completed successfully.")

# List of Excel file paths to process
excel_file_paths = [
    "vulnerability_scores.xlsx",  # Replace with actual paths
    "vulnerability_scores (1).xlsx",
]
column_letter = "A"  # Replace with the desired column letter

# Run the automation
automate_multiple_excel_screenshots(excel_file_paths, column_letter)
