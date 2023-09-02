import win32com.client as win32

# Open Excel and the workbook
excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(r"C:\Users\Rasim\Desktop\quantity_days.xlsb")

# Get the worksheet and range where you want to send the data
worksheet = workbook.Worksheets('Sheet1')
# range = worksheet.Range('A1')
#
# # Send the data to the range
# range.Value = 'Hello, world!'

# Run a VBA macro
excel.Application.Run('quantity_days')

# Save and close the workbook
workbook.Save()
workbook.Close()

# Quit Excel
excel.Quit()
