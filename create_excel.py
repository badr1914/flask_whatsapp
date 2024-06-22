import pandas as pd
from openpyxl import Workbook

file_path = 'messages.xlsx'
sheet_name = 'Sheet1'

# Create a new workbook and worksheet
book = Workbook()
sheet = book.active
sheet.title = sheet_name

# Create initial DataFrame with headers
df = pd.DataFrame(columns=['Sender', 'Message'])

# Write the DataFrame to the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    writer.book = book
    df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"{file_path} created with headers.")
