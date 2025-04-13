using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static NotifyMethods.NotifyMessage;

namespace Grid2Spreadsheet
{
    public static partial class GridSpreadsheet
    {
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
            if (dgv == null)
                throw new ArgumentNullException(nameof(dgv));

            if (string.IsNullOrWhiteSpace(filename))
                return false;

            // Ensure directory exists
            var dir = Path.GetDirectoryName(filename);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            // Figure out which columns to export
            var visibleColumns = dgv.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible)
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            if (visibleColumns.Count == 0)
                return false;

            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add(
                        string.IsNullOrWhiteSpace(worksheetname) ? "Sheet1" : worksheetname);

                    int currentRow = 1;
                    int currentCol = 1;

                    // Write header row
                    if (header)
                    {
                        foreach (var col in visibleColumns)
                        {
                            ws.Cell(currentRow, currentCol)
                              .Value = col.HeaderText;
                            currentCol++;
                        }

                        // Apply bold style to header row (optional)
                        ws.Range(1, 1, 1, visibleColumns.Count)
                          .Style.Font.SetBold();

                        // Apply autofilter to header row
                        ws.Range(1, 1, 1, visibleColumns.Count)
                          .SetAutoFilter();

                        currentRow++;
                    }

                    // Write data rows
                    foreach (DataGridViewRow dgvRow in dgv.Rows)
                    {
                        // skip the 'new row' at bottom
                        if (dgvRow.IsNewRow)
                            continue;

                        currentCol = 1;
                        foreach (var col in visibleColumns)
                        {
                            var cell = ws.Cell(currentRow, currentCol);

                            // grab the cell value; you might want to format dates/numbers etc.
                            var raw = dgvRow.Cells[col.Index].Value;

                            if (raw == null || raw == DBNull.Value)
                            {
                                //cell.Value = string.Empty;
                            }
                            else
                            {
                                // Handle different types explicitly
                                switch (raw)
                                {
                                    case DateTime dt:
                                        cell.Value = dt;
                                        break;

                                    case int i:
                                        cell.Value = i;
                                        break;

                                    case long l:
                                        cell.Value = l;
                                        break;

                                    case double d:
                                        cell.Value = d;
                                        break;

                                    case decimal dec:
                                        cell.Value = (double)dec; // Excel prefers double over decimal
                                        break;

                                    case float f:
                                        cell.Value = (double)f; // Excel prefers double over float
                                        break;

                                    case bool b:
                                        cell.Value = b;
                                        break;

                                    default:
                                        cell.Value = raw.ToString();  // Fallback for other types
                                        break;
                                }
                            }

                            currentCol++;
                        }
                        currentRow++;
                    }

                    // Optionally, you can adjust column widths to fit content:
                    ws.Columns(1, visibleColumns.Count).AdjustToContents();

                    wb.SaveAs(filename);
                    Notify($"Saved to file {filename}");
                }
                return true;
            }
            catch
            {
                // you could log the exception if you want
                return false;
            }
        }
    }
}