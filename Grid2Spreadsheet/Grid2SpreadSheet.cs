using MiniExcelLibs;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static LogicMethods.LogicMethods;
using static NotifyMethods.NotifyMessage;
using Excel = Microsoft.Office.Interop.Excel; //Clear up ambiguity with Application et al

namespace Grid2Spreadsheet
{
    public class GridSpreadsheet
    {
        public static void Grid2Excel(DataGridView dgv, string filename, bool _1 = true)
        //saveFile included for backwards comatibility, not used
        {
            if (string.IsNullOrEmpty(filename))
            {
                Grid2Excel(dgv);
                return;
            }
            var data = new List<Dictionary<string, object>>();
            int batchSize = 1000; // Process 1000 rows at a time
            int totalRows = dgv.Rows.Count;

            using (var stream = File.Create(filename))
            {
                /*var headerRow = new Dictionary<string, object>();
                if (header)
                {
                    foreach (DataGridViewColumn column in dgv.Columns)
                    {
                        if (column.Visible) // Include only visible columns
                        {
                            headerRow[column.HeaderText] = column.HeaderText; // Use the HeaderText as the header in Excel
                        }
                    }
                    data.Add(headerRow);
                }

                int start = 0;
                if (header) start = 1;*/
                for (int start = 0; start < totalRows; start += batchSize)
                {

                    for (int i = start; i < Math.Min(start + batchSize, totalRows); i++)
                    {
                        //if (header && start == 0) { continue; }
                        var row = dgv.Rows[i];
                        if (!row.IsNewRow)
                        {
                            var rowData = new Dictionary<string, object>();
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (dgv.Columns[cell.ColumnIndex].Visible)
                                {
                                    rowData[dgv.Columns[cell.ColumnIndex].HeaderText] = cell.Value;
                                }
                            }
                            data.Add(rowData);
                        }
                    }

                    // Append to the Excel file in batches
                    try
                    {
                        MiniExcel.SaveAs(stream, data);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show($"Error: {ex.Message}");
                    }

                    data.Clear();
                }
            }
            Notify($"Copied to Excel file {filename}");
        }

        public static void Grid2Excel(DataGridView dgv) //, string filename = "", bool saveFile = false)
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
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "JobCheck";
            // storing header part in Excel  
            for (int i = 0, k = 1; i < dgv.Columns.Count; i++)
            {
                if (dgv.Columns[i].Visible == true)
                {
                    worksheet.Cells[1, k] = dgv.Columns[i].HeaderText;
                    k++;
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
            /*if (saveFile && IsTrue(filename))
            {
                // save the application  
                workbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //Exit from the application  
                app.Quit();
            }
            else*/
            {
                //If not saving to file then show the worksheet
                app.Visible = true;
            }
            Notify("Transfered to Excel");
        }
    }
}
