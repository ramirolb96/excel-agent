import calendar
import sys
from datetime import datetime

from openpyxl import load_workbook


def add_next_month(file_path):
    print(f"Loading {file_path}...")
    try:
        wb = load_workbook(file_path)
    except PermissionError:
        print(
            "❌ Error: The file is currently open. Please close Excel/Browser and try again."
        )
        return

    # 1. Determine "Today"
    now = datetime.now()
    current_month_name = now.strftime("%B")  # e.g., "February"

    # 2. Logic: Newest Month is always at index 0 (Leftmost tab)
    latest_sheet = wb.worksheets[0]
    print(f"Most recent tab found: '{latest_sheet.title}'")

    # 3. Determine the Target Month
    if current_month_name == latest_sheet.title:
        # If the first tab is February, we need March
        # Calculate next month using simple math
        next_month_index = (now.month % 12) + 1
        target_month_name = calendar.month_name[next_month_index]
        print(
            f"Current month ({current_month_name}) is up to date. Creating next month: '{target_month_name}'..."
        )
    else:
        # If the first tab is OLD (e.g. January), catch up to Current Month (February)
        target_month_name = current_month_name
        print(
            f"Latest tab is old. Catching up to current month: '{target_month_name}'..."
        )

    if target_month_name in wb.sheetnames:
        print(f"⚠️  Stop: Sheet '{target_month_name}' already exists.")
        return

    # 4. Copy the LEFTMOST sheet (Index 0)
    new_sheet = wb.copy_worksheet(latest_sheet)
    new_sheet.title = target_month_name

    # 5. Move the new sheet to the front (Index 0)
    # By default, copy_worksheet puts it at the end. We move it to the start.
    wb.move_sheet(new_sheet, offset=-(len(wb.sheetnames) - 1))

    # 6. Clear the numbers
    print(f"Clearing old values for {target_month_name}...")
    for row in new_sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue

            # Keep Formulas & Text Labels
            if isinstance(cell.value, str):
                continue

            # Delete Numbers (Expenses)
            if isinstance(cell.value, (int, float)):
                cell.value = None

    # 7. Save
    try:
        wb.save(file_path)
        print(
            f"✅ Success! Created tab '{target_month_name}' at the start of {file_path}"
        )
    except PermissionError:
        print("❌ Error: Could not save. Please close the file in Excel and try again.")


if __name__ == "__main__":
    # Your path
    default_path = "/Users/ramirolb/Library/CloudStorage/OneDrive-Personal/Excel Documents/monthly-expenses.xlsx"
    target_file = sys.argv[1] if len(sys.argv) > 1 else default_path
    add_next_month(target_file)
