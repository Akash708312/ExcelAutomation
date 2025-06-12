

from openpyxl import Workbook
from win32com.client import Dispatch
import os

# 1:Create a new Excel workbook
workbook = Workbook()
sheet = workbook.active

# 2:Add some data
sheet['A1'] = 'Hello'
sheet['B1'] = 'YouTube'

# cell=sheet['A1']
# cell.value="Hello"
# print(cell.value)  can be verified

# 3:Save the Excel file to current directory
file_path = os.path.abspath("hello_youtube.xlsx")
workbook.save(file_path)

# 4:Print the full path on the terminal
print("✅ Excel file saved at:", file_path)

# 5:Launch Excel and open the file
x1 = Dispatch('Excel.Application')
x1.Visible = True  # Show Excel window
x1.Workbooks.Open(file_path)
