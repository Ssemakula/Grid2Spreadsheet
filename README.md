# Grid2Spreadsheet #
DataGridView Extensions

## Description ##
Methods to help in translating DataGridView to spreadsheets
Methods either to open DataGridView in Excel
or save to Excel filename.

Open in Excel is slow

## Methods ##
DataGridView.Grid2Excel([filename = "", header = true, worksheetname = "Sheet1"])

bool IsExcelInstalled()

string GetExcelColumnName(columnIndex)

## Helper Methods ##
DataGridView.OpenInExcel([header = true, worksheetname = "Sheet1"])

DataGridView.Save2ExcelFile(filename, [header = true, worksheetname = "Sheet1"])


 # Licence #
 Licensed under [MIT Licence](https://opensource.org/license/mit)
 
  
