# Import Module 
import win32com.client
import os
# Open Microsoft Excel 
excel = win32com.client.Dispatch("Excel.Application") 
excel.Visible = False
# Read Excel File 
quotation_id = 24
excel_path = os.path.join("quotations", f"{quotation_id}.xlsx")
excel_path = os.path.abspath(excel_path)
pdf_path = os.path.join("quotations", f"{quotation_id}.pdf")
pdf_path = os.path.abspath(pdf_path)

sheets = excel.Workbooks.Open(excel_path) 
work_sheets = sheets.Worksheets[0] 
  
# Convert into PDF File 
work_sheets.ExportAsFixedFormat(0,pdf_path) 