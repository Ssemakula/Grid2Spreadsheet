using System;
using System.Windows.Forms;
using static NotifyMethods.NotifyMessage;
using Excel = Microsoft.Office.Interop.Excel; //Clear up ambiguity with Application et al

namespace Grid2Spreadsheet
{
    public static partial class GridSpreadsheet
    {
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
    }
}