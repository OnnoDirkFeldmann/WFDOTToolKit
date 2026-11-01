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
                dialog.Title = "WFDOT Excel-Datei auswählen";
                dialog.Filter = "Excel-Arbeitsmappe (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|Alle Dateien (*.*)|*.*";
                Log("Datei-Dialog geöffnet, warte auf Auswahl ...");
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    Log("Abgebrochen (keine Datei ausgewählt).");
                    return;
                }

                Log($"Ausgewählte Datei: {dialog.FileName}");
                Process(dialog.FileName);
            }
        }

        public static void Process(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("Excel-Datei nicht gefunden.", excelPath);
            }

            var backupPath = Path.Combine(
                Path.GetDirectoryName(excelPath),
                Path.GetFileNameWithoutExtension(excelPath) + "_backup_nummerierung_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(excelPath));
            Log($"Erstelle Backup: {backupPath}");
            File.Copy(excelPath, backupPath, true);

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                Log("Starte Excel ...");
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                Log("Öffne Excel-Datei (beschreibbar) ...");
                xlWorkbook = xlApp.Workbooks.Open(excelPath);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                Log($"Verwende Arbeitsblatt: \"{xlWorksheet.Name}\"");
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                Log($"Erkannter Bereich: {rowCount - 1} Datenzeilen, {colCount} Spalten");

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
                    throw new InvalidOperationException("Spalte 'Ostfriesisch' nicht gefunden.");
                }
                if (!headerToColumn.TryGetValue("Nummer", out int numCol))
                {
                    throw new InvalidOperationException("Spalte 'Nummer' nicht gefunden.");
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
                    Log("Keine unnummerierten Duplikate gefunden.");
                    return;
                }

                Log($"Gefunden: {duplicatesToProcess.Count} Ostfriesisch-Einträge mit Duplikaten ohne vollständige Nummerierung.");
                Console.WriteLine();
                Console.WriteLine("Für jeden Duplikat-Eintrag kann eine Nummer vergeben werden.");
                Console.WriteLine("Wenn bereits eine Nummer vorhanden ist (z. B. 1), kann diese bestätigt/übersprungen werden.");
                Console.WriteLine("Antworten: Zahl (z. B. 1,2,3) = Nummer setzen, s = überspringen, a = automatisch durchnummerieren, q = abbrechen");
                Console.WriteLine();

                int updated = 0;
                foreach (var kv in duplicatesToProcess)
                {
                    string key = kv.Key;
                    var rows = kv.Value.OrderBy(r => r.Row).ToList();
                    Console.WriteLine($"=== Duplikate für \"{key}\" ({rows.Count} Einträge) ===");
                    
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
                        Console.WriteLine($"  Zeile {rowInfo.Row}: Ostfriesisch=\"{rowInfo.Ostfriesisch}\", Nummer=\"{(string.IsNullOrEmpty(rowInfo.Nummer) || rowInfo.Nummer == "-" ? "(leer)" : rowInfo.Nummer)}\"");
                    }
                    Console.WriteLine();

                    bool groupAuto = false;
                    for (int idx = 0; idx < rows.Count; idx++)
                    {
                        var rowInfo = rows[idx];
                        bool hasNumber = !string.IsNullOrEmpty(rowInfo.Nummer) && rowInfo.Nummer != "-" && !rowInfo.Nummer.Equals("0", StringComparison.OrdinalIgnoreCase);
                        if (hasNumber && !groupAuto)
                        {
                            Console.Write($"  Zeile {rowInfo.Row} hat bereits Nummer \"{rowInfo.Nummer}\" - beibehalten? (j/n/q): ");
                            string ans = (Console.ReadLine() ?? string.Empty).Trim().ToLowerInvariant();
                            if (ans == "q") { Log("Abbruch durch Benutzer."); goto SaveAndExit; }
                            if (ans == "n" || ans == "nein") { hasNumber = false; }
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
                                Console.WriteLine($"    -> Zeile {rowInfo.Row}: Nummer auf {numToSet} gesetzt (auto)");
                                continue;
                            }

                            Console.Write($"  Nummer für Zeile {rowInfo.Row} festlegen (Enter = {nextNum}, 's' skip, 'a' auto für Rest, 'q' abbrechen): ");
                            string input = (Console.ReadLine() ?? string.Empty).Trim().ToLowerInvariant();
                            if (input == "q") { Log("Abbruch durch Benutzer."); goto SaveAndExit; }
                            if (input == "s" || input == "skip") { continue; }
                            if (input == "a" || input == "auto") 
                            { 
                                groupAuto = true;
                                string numToSet = nextNum.ToString();
                                nextNum++;
                                SetCell(xlWorksheet, rowInfo.Row, numCol, numToSet);
                                rowInfo.Nummer = numToSet;
                                updated++;
                                Console.WriteLine($"    -> Zeile {rowInfo.Row}: Nummer auf {numToSet} gesetzt (auto)");
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
                                Console.WriteLine($"    -> Zeile {rowInfo.Row}: Nummer auf {setNum} gesetzt");
                            }
                            else
                            {
                                Console.WriteLine("    -> Ungültig, übersprungen");
                            }
                        }
                    }
                    Console.WriteLine();
                }

                SaveAndExit:
                if (updated > 0)
                {
                    Log($"Speichere {updated} Änderungen ...");
                    xlWorkbook.Save();
                    Log("Excel-Datei gespeichert.");
                }
                else
                {
                    Log("Keine Änderungen.");
                }
            }
            catch (Exception ex)
            {
                Log($"=== FEHLER: {ex.Message} ===");
                throw;
            }
            finally
            {
                Log("Schließe Excel und räume auf ...");
                if (xlWorkbook != null) { xlWorkbook.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook); }
                if (xlWorksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                if (xlRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                if (xlApp != null) { xlApp.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp); }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Log("Aufgeräumt.");
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
