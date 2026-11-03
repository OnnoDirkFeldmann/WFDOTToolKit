using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WFDOTToolKit
{
    internal static class WFDOTDuplicateNumbering
    {
        private static void Log(string message)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }

        public static void Run()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select WFDOT Excel file";
                dialog.Filter = "Excel workbook (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|All files (*.*)|*.*";
                Log("File dialog opened, waiting for selection ...");
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    Log("Cancelled (no file selected).");
                    return;
                }

                Log($"Selected file: {dialog.FileName}");
                Process(dialog.FileName);
            }
        }

        public static void Process(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("Excel file not found.", excelPath);
            }

            var backupPath = Path.Combine(
                Path.GetDirectoryName(excelPath),
                Path.GetFileNameWithoutExtension(excelPath) + "_backup_nummerierung_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(excelPath));
            Log($"Created backup: {backupPath}");
            File.Copy(excelPath, backupPath, true);

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                Log("Starting Excel ...");
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                Log("Opening Excel file (writable) ...");
                xlWorkbook = xlApp.Workbooks.Open(excelPath);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                Log($"Using worksheet: \"{xlWorksheet.Name}\"");
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                Log($"Detected range: {rowCount - 1} data rows, {colCount} columns");

                var cells = (object[,])xlRange.Value2;
                var formulas = (object[,])xlRange.Formula;

                Func<int, int, bool> isFormula = (row, col) =>
                {
                    if (formulas[row, col] == null)
                    {
                        return false;
                    }
                    return Convert.ToString(formulas[row, col]).StartsWith("=", StringComparison.Ordinal);
                };

                var headerToColumn = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    string header = CellText(cells, 1, c);
                    if (!string.IsNullOrWhiteSpace(header) && !headerToColumn.ContainsKey(header))
                    {
                        headerToColumn[header] = c;
                    }
                }

                if (!headerToColumn.TryGetValue("Ostfriesisch", out int ostCol))
                {
                    throw new InvalidOperationException("Column 'Ostfriesisch' not found.");
                }
                if (!headerToColumn.TryGetValue("Nummer", out int numCol))
                {
                    throw new InvalidOperationException("Column 'Nummer' not found.");
                }

                // Group duplicates by Ostfriesisch
                var groups = new Dictionary<string, List<RowInfo>>(StringComparer.OrdinalIgnoreCase);
                for (int r = 2; r <= rowCount; r++)
                {
                    if (isFormula(r, ostCol))
                    {
                        continue;
                    }
                    string ost = CellText(cells, r, ostCol).Trim();
                    if (string.IsNullOrEmpty(ost))
                    {
                        continue;
                    }
                    string num = CellText(cells, r, numCol).Trim();
                    if (!groups.TryGetValue(ost, out var list))
                    {
                        list = new List<RowInfo>();
                        groups[ost] = list;
                    }
                    list.Add(new RowInfo
                    {
                        Row = r,
                        Ostfriesisch = ost,
                        Nummer = num
                    });
                }

                var duplicatesToProcess = groups
                    .Where(g => g.Value.Count > 1)
                    .Where(g => g.Value.Any(x => string.IsNullOrEmpty(x.Nummer) || x.Nummer == "-" || x.Nummer.Equals("0", StringComparison.OrdinalIgnoreCase)))
                    .OrderBy(g => g.Key)
                    .ToList();

                if (duplicatesToProcess.Count == 0)
                {
                    Log("No unnumbered duplicates found.");
                    return;
                }

                Log($"Found: {duplicatesToProcess.Count} Ostfriesisch entries with duplicates lacking a complete number.");
                Console.WriteLine();
                Console.WriteLine("A number can be assigned to each duplicate entry.");
                Console.WriteLine("If a number is already present (e.g. 1), it can be confirmed/skipped.");
                Console.WriteLine("Answers: number (e.g. 1,2,3) = set number, s = skip, a = automatically number, q = cancel");
                Console.WriteLine();

                int updated = 0;
                foreach (var kv in duplicatesToProcess)
                {
                    string key = kv.Key;
                    var rows = kv.Value.OrderBy(r => r.Row).ToList();
                    Console.WriteLine($"=== Duplicates for \"{key}\" ({rows.Count} entries) ===");
                    
                    // Find next available number for this group
                    var existingNums = rows
                        .Where(r => !string.IsNullOrEmpty(r.Nummer) && r.Nummer != "-" && !r.Nummer.Equals("0", StringComparison.OrdinalIgnoreCase))
                        .Select(r => r.Nummer)
                        .ToList();
                    
                    int nextNum = 1;
                    foreach (var n in existingNums)
                    {
                        if (int.TryParse(n, out int parsed) && parsed >= nextNum)
                        {
                            nextNum = parsed + 1;
                        }
                    }

                    foreach (var rowInfo in rows)
                    {
                        Console.WriteLine($"  Row {rowInfo.Row}: Ostfriesisch=\"{rowInfo.Ostfriesisch}\", Nummer=\"{(string.IsNullOrEmpty(rowInfo.Nummer) || rowInfo.Nummer == "-" ? "(empty)" : rowInfo.Nummer)}\"");
                    }
                    Console.WriteLine();

                    bool groupAuto = false;
                    for (int idx = 0; idx < rows.Count; idx++)
                    {
                        var rowInfo = rows[idx];
                        bool hasNumber = !string.IsNullOrEmpty(rowInfo.Nummer) && rowInfo.Nummer != "-" && !rowInfo.Nummer.Equals("0", StringComparison.OrdinalIgnoreCase);
                        if (hasNumber && !groupAuto)
                        {
                            Console.Write($"  Row {rowInfo.Row} already has number \"{rowInfo.Nummer}\" - keep? (j/n/q): ");
                            string ans = (Console.ReadLine() ?? string.Empty).Trim().ToLowerInvariant();
                            if (ans == "q") { Log("User cancelled."); goto SaveAndExit; }
                            if (ans == "n") { hasNumber = false; }
                            else { continue; } // keep existing
                        }

                        if (!hasNumber || groupAuto)
                        {
                            if (groupAuto)
                            {
                                string numToSet = nextNum.ToString();
                                nextNum++;
                                SetCell(xlWorksheet, rowInfo.Row, numCol, numToSet);
                                rowInfo.Nummer = numToSet;
                                updated++;
                                Console.WriteLine($"    -> Row {rowInfo.Row}: number set to {numToSet} (auto)");
                                continue;
                            }

                            Console.Write($"  Enter a number for row {rowInfo.Row} (Enter = {nextNum}, 's' skip, 'a' auto for the rest, 'q' cancel): ");
                            string input = (Console.ReadLine() ?? string.Empty).Trim().ToLowerInvariant();
                            if (input == "q") { Log("User cancelled."); goto SaveAndExit; }
                            if (input == "s" || input == "skip") { continue; }
                            if (input == "a" || input == "auto") 
                            { 
                                groupAuto = true;
                                string numToSet = nextNum.ToString();
                                nextNum++;
                                SetCell(xlWorksheet, rowInfo.Row, numCol, numToSet);
                                rowInfo.Nummer = numToSet;
                                updated++;
                                Console.WriteLine($"    -> Row {rowInfo.Row}: number set to {numToSet} (auto)");
                                continue; 
                            }
                            if (string.IsNullOrEmpty(input))
                            {
                                input = nextNum.ToString();
                            }
                            if (int.TryParse(input, out int setNum))
                            {
                                SetCell(xlWorksheet, rowInfo.Row, numCol, setNum.ToString());
                                rowInfo.Nummer = setNum.ToString();
                                if (setNum >= nextNum) nextNum = setNum + 1;
                                updated++;
                                Console.WriteLine($"    -> Row {rowInfo.Row}: number set to {setNum}");
                            }
                            else
                            {
                                Console.WriteLine("    -> Invalid, skipped");
                            }
                        }
                    }
                    Console.WriteLine();
                }

                SaveAndExit:
                if (updated > 0)
                {
                    Log($"Saving {updated} changes ...");
                    xlWorkbook.Save();
                    Log("Excel file saved.");
                }
                else
                {
                    Log("No changes.");
                }
            }
            catch (Exception ex)
            {
                Log($"=== ERROR: {ex.Message} ===");
                throw;
            }
            finally
            {
                Log("Closing Excel and cleaning up ...");
                if (xlWorkbook != null) { xlWorkbook.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook); }
                if (xlWorksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                if (xlRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                if (xlApp != null) { xlApp.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp); }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Log("Cleaned up.");
            }
        }

        private static void SetCell(Excel.Worksheet sheet, int row, int col, string value)
        {
            sheet.Cells[row, col] = value;
        }

        private static string CellText(object[,] cells, int row, int column)
        {
            object value = cells[row, column];
            return value == null ? string.Empty : Convert.ToString(value);
        }

        private class RowInfo
        {
            public int Row { get; set; }
            public string Ostfriesisch { get; set; }
            public string Nummer { get; set; }
        }
    }
}
