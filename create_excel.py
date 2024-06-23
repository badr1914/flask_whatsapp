from openpyxl import Workbook

# Create a new Workbook
wb = Workbook()

# Get the active sheet
ws = wb.active

# Rename the active sheet to "sheet1"
ws.title = "sheet1"

# Save the workbook as 'message.xlsx'
wb.save("message.xlsx")

print("Excel file 'message.xlsx' with sheet 'sheet1' has been created.")
