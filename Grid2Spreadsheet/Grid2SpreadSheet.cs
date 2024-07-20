using System;
using System.Windows.Forms;
using static LogicMethods.LogicMethods;
using static NotifyMethods.NotifyMessage;
using Excel = Microsoft.Office.Interop.Excel; //Clear up ambiguity with Application et al

namespace Grid2Spreadsheet
{
    public class GridSpreadsheet
    {
        public static void Grid2Excel(DataGridView dgv, string filename = "", bool saveFile = false)
        {
            // creating Excel Application  
            Excel._Application app = new Excel.Application();
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
            if (saveFile && IsTrue(filename))
            {
                // save the application  
                workbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //Exit from the application  
                app.Quit();
            }
            else
            {
                //If not saving to file then show the worksheet
                app.Visible = true;
            }
            Notify("Transfered to Excel");
        }
    }
}
