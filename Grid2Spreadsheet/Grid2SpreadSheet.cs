using System;
using System.Windows.Forms;
using static NotifyMethods.NotifyMessage;

namespace Grid2Spreadsheet
{
    public static partial class GridSpreadsheet
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

        /// <summary>
        /// Helper to convert column index to Excel letter (e.g., 0 → "A")
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns>Name of column</returns>
        public static string GetExcelColumnName(int columnIndex)
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