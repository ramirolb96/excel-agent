import calendar
import sys
from datetime import datetime

from openpyxl import load_workbook


def add_next_month(file_path):
    print(f"Loading {file_path}...")
    try:
        wb = load_workbook(file_path)
    except PermissionError:
        print("❌ Error: The file is open. Please close Excel/Browser and try again.")
        return

    # 1. Determine "Today"
    now = datetime.now()
    current_month_name = now.strftime("%B")

    # 2. Get the Newest Tab (First one on the left)
    latest_sheet = wb.worksheets[0]

    # 3. Figure out the Target Month
    if current_month_name == latest_sheet.title:
        # If February exists, we create March
        next_month_index = (now.month % 12) + 1
        target_month_name = calendar.month_name[next_month_index]
    else:
        # If February is missing, we create February
        target_month_name = current_month_name

    if target_month_name in wb.sheetnames:
        print(f"⚠️  Stop: Sheet '{target_month_name}' already exists.")
        return

    print(f"Creating '{target_month_name}' from template '{latest_sheet.title}'...")

    # 4. Copy and Move to Front
    new_sheet = wb.copy_worksheet(latest_sheet)
    new_sheet.title = target_month_name
    wb.move_sheet(new_sheet, offset=-(len(wb.sheetnames) - 1))

    # 5. Clear Data (Numbers AND Dates)
    print(f"Clearing expenses and dates for {target_month_name}...")

    for row in new_sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue

            # KEEP: Formulas (starts with =)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue

            # KEEP: Text Labels (e.g. "Week 1", "Car Insurance")
            if isinstance(cell.value, str):
                continue

            # DELETE: Numbers (Prices)
            if isinstance(cell.value, (int, float)):
                cell.value = None

            # DELETE: Dates (This fixes the Due Date issue)
            if isinstance(cell.value, datetime):
                cell.value = None

    # 6. Save
    try:
        wb.save(file_path)
        print(f"✅ Success! Created '{target_month_name}'")
    except PermissionError:
        print("❌ Error: Could not save. Close the file and try again.")


if __name__ == "__main__":
    # Update this path if needed
    default_path = "/Users/ramirolb/Library/CloudStorage/OneDrive-Personal/Excel Documents/monthly-expenses.xlsx"
    target_file = sys.argv[1] if len(sys.argv) > 1 else default_path
    add_next_month(target_file)
