# allintitile-identifier-excel

```
import openpyxl
import tkinter as tk
from tkinter import filedialog

# Create a GUI window to select the Excel file
root = tk.Tk()
root.withdraw()  # Hide the main window

# Ask the user to select an Excel file
file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

if not file_path:
    print("No file selected. Exiting.")
else:
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(file_path)

    # Select the specific sheet within the workbook (replace 'Sheet1' with your sheet name)
    sheet = workbook['Sheet1']

    # Find the next available column for marking rows
    next_column = sheet.max_column + 1

    # Iterate through rows in column D (assuming your data starts from row 2)
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True), start=2):
        cell_value = row[0]

        try:
            # Try to convert the cell value to a float
            cell_value_float = float(cell_value)

            # Check if the value in column D is less than 1
            if cell_value_float < 1:
                # Mark the row with a tick in the next column
                sheet.cell(row=row_number, column=next_column).value = "✔"  # This example uses a checkmark symbol
        except ValueError:
            # Handle non-numeric values (e.g., '#DIV/0!') here
            pass

    # Save the modified Excel file
    modified_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")

    if modified_file_path:
        workbook.save(modified_file_path)
        print("File has been modified and saved as:", modified_file_path)
    else:
        print("No file selected for saving. Exiting.")
```











