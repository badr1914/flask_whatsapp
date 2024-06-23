from openpyxl import Workbook

# Create a new workbook and set the active sheet
wb = Workbook()
ws = wb.active
ws.title = "TestSheet"

# Add headers to the first row
headers = ['Header1', 'Header2']
ws.append(headers)

# Save the workbook
wb.save('test.xlsx')

print("test.xlsx created successfully")
