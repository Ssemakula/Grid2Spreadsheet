using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Numerics;
using static LogicMethods.LogicMethods;
using static NotifyMethods.NotifyMessage;
using Excel = Microsoft.Office.Interop.Excel; //Clear up ambiguity with Application et al
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.CompilerServices;
using System.Globalization;

namespace Grid2Spreadsheet
{
    public static class GridSpreadsheet
    {
        /// <summary>
        /// Saves a DataGridView to Excel file or Excel sheet
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="header">Optional. true (default) if headers are required</param>
        /// <param name="worksheetname">Optional. Defaults to Sheet1</param>
        /// <param name="filename">Optiona. If empty or whitespace will load Excel</param>
        public static void Grid2Excel(this DataGridView dgv, string filename = "", bool header = true, string worksheetname = "Sheet1")
        //saveFile included for backwards comatibility, not used
        {
            //Check if grid is empty
            if (dgv.RowCount == 0) return;

            if (string.IsNullOrEmpty(filename))
            {
                if (!IsExcelInstalled())
                {
                    Notify("Unable to initialise Excel");
                    return;
                }
                dgv.OpenInExcel(header, worksheetname);
            }
            else
            {
                dgv.Save2ExcelFile(filename, header, worksheetname);
            }
        }

        /// <summary>
        /// Opens a DataGridView in Excel
        /// Usually called by Grid2Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="header">Sepcifies whether headers are required. Defaults to true</param>
        /// <param name="worksheetName">Specified Worhsheet name, defaults to Sheet1</param>
        public static void OpenInExcel(this DataGridView dgv, bool header = true, string worksheetName = "Sheet1")
        {
            Excel._Application app;
            // creating Excel Application
            try
            {
                app = new Excel.Application();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("This option requires Excel to be installed and activated");
                return;
            }
            // creating new WorkBook within Excel application
            Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook
            Excel._Worksheet worksheet = null;
            // initially hide the excel sheet behind the program
            app.Visible = false;
            // get the reference of first sheet.
            // By default its name is Sheet1.
            // If this changes we'll need to rewrite this
            // store its reference to worksheet
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet if requested
            if (!string.IsNullOrWhiteSpace(worksheetName))
            {
                worksheet.Name = worksheetName;
            }

            int visibleColumnsCount = 0; // Get number of visible columns
            // storing header part in Excel
            if (header)
            {
                for (int i = 0, k = 1; i < dgv.Columns.Count; i++)
                {
                    if (dgv.Columns[i].Visible == true)
                    {
                        worksheet.Cells[1, k] = dgv.Columns[i].HeaderText;
                        k++;
                        visibleColumnsCount = k - 1;
                    }
                }
            }

            // storing Each row and column value to excel sheet
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                for (int j = 0, k = 1; j < dgv.Columns.Count; j++)
                {
                    if (dgv.Columns[j].Visible == true)
                    {
                        worksheet.Cells[i + 2, k] = dgv.Rows[i].Cells[j].Value; //.ToString();
                        k++;
                    }
                }
            }

            //Apply auto filter
            if (header && visibleColumnsCount > 0)
            {
                // Define the range: Header (Row 1) + Data (Row 2 onwards)
                Excel.Range headerRange = worksheet.Range[
                    worksheet.Cells[1, 1],
                    worksheet.Cells[dgv.Rows.Count + 1, visibleColumnsCount] // +1 to include header
                ];

                headerRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            }

            app.Visible = true;
            Notify("Transfered to Excel");
        }

        /// <summary>
        /// Saves DataGridView to Excel file
        /// (<paramref name="filename"/>).
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="filename">Filename (full or relative path)</param>
        /// <param name="header">specifies if headers are required. Defaults to true</param>
        /// <param name="worksheetname">Worksheet name. Defaults to Sheet1</param>
        /// <returns></returns>
        public static bool Save2ExcelFile(this DataGridView dgv, string filename, bool header = true, string worksheetname = "Sheet1")
        {
            if (string.IsNullOrWhiteSpace(filename))
            {
                return false;
            }

            string wsn = worksheetname;

            if (string.IsNullOrWhiteSpace(wsn))
            {
                wsn = "Sheet1";
            }

            var visibleColumns = dgv.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible)
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = wsn };
                sheets.Append(sheet);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Write headers
                if (header)
                {
                    Row headerRow = new Row();
                    foreach (var column in visibleColumns)
                    {
                        headerRow.Append(new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.HeaderText)
                        });
                    }
                    sheetData.AppendChild(headerRow);
                }

                // Write data rows
                foreach (DataGridViewRow dgvRow in dgv.Rows)
                {
                    if (dgvRow.Visible && !dgvRow.IsNewRow)
                    {
                        Row dataRow = new Row();
                        foreach (var column in visibleColumns)
                        {
                            var cell = dgvRow.Cells[column.Index];
                            object cellValue = cell.Value;

                            Cell excelCell = new Cell(); // Default: empty cell

                            if (cellValue != null)
                            {
                                if (IsBoolean(cellValue))
                                {
                                    excelCell.DataType = CellValues.Boolean;
                                    excelCell.CellValue = new CellValue((bool)cellValue ? "1" : "0");
                                }
                                else if (IsDateTime(cellValue))
                                {
                                    double oaDate = ((DateTime)cellValue).ToOADate();
                                    excelCell.DataType = CellValues.Number;
                                    excelCell.CellValue = new CellValue(oaDate.ToString(CultureInfo.InvariantCulture));
                                }
                                else if (IsNumeric(cellValue))
                                {
                                    excelCell.DataType = CellValues.Number;
                                    excelCell.CellValue = new CellValue(
                                        Convert.ToString(cellValue, CultureInfo.InvariantCulture)
                                    );
                                }
                                else
                                {
                                    // Treat non-null, non-special types as strings
                                    excelCell.DataType = CellValues.String;
                                    excelCell.CellValue = new CellValue(cellValue.ToString());
                                }
                            }
                            // If cellValue is null, leave the cell empty (no DataType/CellValue)

                            dataRow.Append(excelCell);
                        }
                        sheetData.AppendChild(dataRow);
                    }
                }
                // --- Add AutoFilter to Headers ---
                if (header && visibleColumns.Count > 0)
                {
                    int totalRows = sheetData.Elements<Row>().Count(); // Includes header + data
                    string startCol = GetExcelColumnName(0); // "A"
                    string endCol = GetExcelColumnName(visibleColumns.Count - 1); // e.g., "C"

                    worksheetPart.Worksheet.Append(new AutoFilter()
                    {
                        Reference = $"{startCol}1:{endCol}{totalRows}" // e.g., "A1:C10"
                    });
                }

                workbookPart.Workbook.Save();
            }

            Notify($"Saved to Excel File {filename}");
            return true;
        }

        /// <summary>
        /// Checks whether the systems contains an activated installation of Microsoft Excel
        /// </summary>
        /// <returns></returns>
        public static bool IsExcelInstalled()
        {
            bool result;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                result = excelType != null;
            }
            catch
            {
                result = false;
            }

            return result;
        }

        // Check if a value is numeric (int, double, decimal, etc.)
        private static bool IsNumeric(object value)
        {
            if (value == null) return false;
            switch (Type.GetTypeCode(value.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Decimal:
                case TypeCode.Single:
                case TypeCode.Double:
                    return true;

                default:
                    return false;
            }
        }

        // Check if a value is a boolean
        private static bool IsBoolean(object value)
        {
            return value is bool;
        }

        // Check if a value is a DateTime
        private static bool IsDateTime(object value)
        {
            return value is DateTime;
        }

        // Helper to convert column index to Excel letter (e.g., 0 → "A")
        private static string GetExcelColumnName(int columnIndex)
        {
            string columnName = "";
            while (columnIndex >= 0)
            {
                int remainder = columnIndex % 26;
                columnName = (char)(65 + remainder) + columnName;
                columnIndex = (columnIndex / 26) - 1;
            }
            return columnName;
        }
    }
}