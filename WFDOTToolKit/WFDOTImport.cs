using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WFDOTToolKit
{
    /// <summary>
    /// Fast replacement for the former manual import chain
    /// (Excel copy -> delete Projekt/Operation/Operation2 -> save as Unicode text ->
    ///  Notepad++ macro -> SQLiteStudio update script).
    ///
    /// The original Excel file is opened read-only and is never modified.
    /// </summary>
    internal static class WFDOTImport
    {
        // Columns that exist in the worksheet but must NOT be imported.
        private static readonly HashSet<string> IgnoredColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Projekt", "Operation", "Operation2"
        };

        // Column order as stored in the WB table (everything except the auto ID).
        private static readonly string[] DbColumns =
        {
            "Ostfriesisch", "Deutsch", "Artikel", "Wortart", "Plural", "Genus",
            "Komparation", "Konjugation", "Nebenformen", "Standardform", "Zuordnung",
            "Nummer", "Etymologie", "Kommentar", "Rezept", "Englisch"
        };

        // FTS source columns (same as the fts4 definition in the DB).
        private static readonly string[] FtsColumns =
        {
            "ID", "Ostfriesisch", "Deutsch", "Artikel", "Wortart", "Plural", "Genus",
            "Komparation", "Konjugation", "Nebenformen", "Standardform", "Zuordnung",
            "Nummer", "Englisch"
        };

        // The Ostfriesisch column stores clitics with a leading apostrophe ('k, 't, 's, 'n).
        // When such a form is entered without the apostrophe (k heb..., t is maal, lone k)
        // it gets one. Other columns are never touched, so e.g. English sentences that
        // merely start with a one-letter word stay intact.
        private static readonly Regex CliticRule = new Regex(@"^([knst])(\s|$)");

        /// <summary>Writes a timestamped progress message to the console.</summary>
        private static void Log(string message)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }

        public static void Import()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "WFDOT Excel-Datei auswählen";
                dialog.Filter = "Excel-Arbeitsmappe (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|Alle Dateien (*.*)|*.*";
                Log("Datei-Dialog geöffnet, warte auf Auswahl ...");
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    Log("Import abgebrochen (keine Datei ausgewählt).");
                    return;
                }

                Log($"Ausgewählte Datei: {dialog.FileName}");
                Import(dialog.FileName, WfdotPath);
            }
        }

        /// <summary>Absolute path of the database that lives next to the executable's source folder.</summary>
        private static string WfdotPath =>
            @"D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";

        public static void Import(string excelPath, string sqlitePath)
        {
            var stopwatch = Stopwatch.StartNew();
            Log("=== Import gestartet ===");
            Log($"Excel-Datei : {excelPath}");
            Log($"Datenbank   : {sqlitePath}");

            // Operate on a copy so the original workbook stays untouched even if Excel offers to save.
            var workbookCopy = Path.Combine(Path.GetTempPath(), "wfdot_import_" + Guid.NewGuid().ToString("N") + Path.GetExtension(excelPath));
            Log("Erstelle Arbeitskopie der Excel-Datei (Original bleibt unverändert) ...");
            File.Copy(excelPath, workbookCopy, true);
            Log($"Arbeitskopie: {workbookCopy}");

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                Log("Starte Excel ...");
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                Log("Öffne Arbeitskopie in Excel (schreibgeschützt) ...");
                xlWorkbook = xlApp.Workbooks.Open(workbookCopy, ReadOnly: true);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                Log($"Verwende Arbeitsblatt: \"{xlWorksheet.Name}\"");
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                Log($"Erkannter Bereich: {rowCount - 1} Datenzeilen, {colCount} Spalten");

                // Read the entire used range in ONE COM call instead of one call per cell.
                // This is the difference between minutes and seconds for large sheets.
                Log("Lese den gesamten Bereich auf einmal aus Excel (kann bei großen Dateien kurz dauern) ...");
                var readWatch = Stopwatch.StartNew();
                var cells = (object[,])xlRange.Value2;
                Log($"Bereich ausgelesen in {readWatch.Elapsed.TotalSeconds:F1} s.");

                // Map header names -> worksheet column index, honouring the ignore list.
                Log("Lese Spaltenüberschriften ...");
                var headerToColumn = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    string header = CellText(cells, 1, c);
                    if (string.IsNullOrWhiteSpace(header) || IgnoredColumns.Contains(header))
                    {
                        if (IgnoredColumns.Contains(header))
                        {
                            Log($"  Spalte {c} \"{header}\" wird übersprungen (nicht importieren).");
                        }
                        continue;
                    }
                    if (!headerToColumn.ContainsKey(header))
                    {
                        headerToColumn[header] = c;
                    }
                }

                var missing = new List<string>();
                foreach (var column in DbColumns)
                {
                    if (!headerToColumn.ContainsKey(column))
                    {
                        missing.Add(column);
                    }
                }
                if (missing.Count > 0)
                {
                    throw new InvalidOperationException(
                        "In der Excel-Datei fehlen folgende Spalten: " + string.Join(", ", missing));
                }
                Log($"Alle {DbColumns.Length} benötigten Spalten gefunden.");

                Log("Öffne Datenbank ...");
                using (var connection = new SQLiteConnection("Data Source=" + sqlitePath))
                {
                    connection.Open();

                    using (var transaction = connection.BeginTransaction())
                    {
                        Log("Leere Tabelle WB ...");
                        Exec(connection, transaction,
                            "CREATE TABLE IF NOT EXISTS WB (ID INTEGER NOT NULL UNIQUE, " +
                            "Ostfriesisch TEXT NOT NULL, Deutsch TEXT NOT NULL, Artikel TEXT NOT NULL, " +
                            "Wortart TEXT NOT NULL, Plural TEXT NOT NULL, Genus TEXT NOT NULL, " +
                            "Komparation TEXT NOT NULL, Konjugation TEXT NOT NULL, Nebenformen TEXT NOT NULL, " +
                            "Standardform TEXT NOT NULL, Zuordnung TEXT NOT NULL, Nummer TEXT NOT NULL, " +
                            "Etymologie TEXT NOT NULL, Kommentar TEXT NOT NULL, Rezept TEXT NOT NULL, " +
                            "Englisch TEXT NOT NULL, PRIMARY KEY (ID));");
                        Exec(connection, transaction, "DELETE FROM WB;");
                        Log("Tabelle WB geleert.");

                        Log("Schreibe Datenzeilen ...");
                        int imported = 0;
                        int skipped = 0;
                        int corrected = 0;
                        int trimmed = 0;

                        using (var insert = connection.CreateCommand())
                        {
                            insert.Transaction = transaction;
                            var columnList = string.Join(", ", DbColumns);
                            var parameterNames = new string[DbColumns.Length];
                            for (int i = 0; i < DbColumns.Length; i++)
                            {
                                parameterNames[i] = "@p" + i;
                            }
                            var parameterList = string.Join(", ", parameterNames);
                            insert.CommandText = $"INSERT INTO WB ({columnList}) VALUES ({parameterList});";
                            var parameters = new SQLiteParameter[DbColumns.Length];
                            for (int i = 0; i < DbColumns.Length; i++)
                            {
                                parameters[i] = insert.Parameters.Add(parameterNames[i], System.Data.DbType.String);
                            }

                            for (int r = 2; r <= rowCount; r++)
                            {
                                bool empty = true;
                                for (int i = 0; i < DbColumns.Length; i++)
                                {
                                    string columnName = DbColumns[i];
                                    string original = CellText(cells, r, headerToColumn[columnName]);
                                    string value = Clean(columnName, original, out bool apostropheAdded, out bool whitespaceTrimmed);
                                    if (whitespaceTrimmed)
                                    {
                                        trimmed++;
                                    }
                                    if (apostropheAdded)
                                    {
                                        corrected++;
                                        Log($"  Korrektur Zeile {r}, Spalte \"{columnName}\": \"{original}\" -> \"{value}\"");
                                    }
                                    parameters[i].Value = value;
                                    if (value.Length > 0)
                                    {
                                        empty = false;
                                    }
                                }

                                if (empty)
                                {
                                    skipped++;
                                    Log($"  Leere Zeile {r} übersprungen.");
                                    continue; // skip fully empty rows
                                }

                                insert.ExecuteNonQuery();
                                imported++;

                                if (imported % 10000 == 0)
                                {
                                    Log($"  ... {imported} Zeilen geschrieben");
                                }
                            }
                        }

                        Log($"Datenzeilen geschrieben: {imported}" +
                            (skipped > 0 ? $", übersprungen: {skipped}" : "") +
                            (corrected > 0 ? $", Apostroph-Korrekturen: {corrected}" : "") +
                            (trimmed > 0 ? $", Leerzeichen am Ende entfernt: {trimmed}" : ""));
                        Log("Übernehme Änderungen (Commit) ...");
                        transaction.Commit();
                        Log("Änderungen übernommen.");
                    }

                    Log("Baue Suchindex (WBFTS) neu auf ...");
                    RebuildFts(connection);
                    Log("Suchindex neu aufgebaut.");
                    Log("Datenbank geschlossen.");
                }

                Log($"Datenbank fertig in {stopwatch.Elapsed.TotalSeconds:F1} s.");
            }
            catch (Exception ex)
            {
                Log($"=== FEHLER: {ex.Message} ===");
                throw;
            }
            finally
            {
                Log("Schließe Excel und räume auf ...");
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                }
                if (xlWorksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                }
                if (xlRange != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                try { File.Delete(workbookCopy); } catch { /* best effort */ }
                Log("Aufgeräumt.");
            }

            stopwatch.Stop();
            Log($"=== Import erfolgreich abgeschlossen in {stopwatch.Elapsed.TotalSeconds:F1} s ===");
        }

        /// <summary>
        /// Cleans the Excel file itself (in place):
        ///   - removes trailing whitespace from every cell,
        ///   - fixes Ostfriesisch clitics (k -> 'k, "t is maal" -> "'t is maal").
        /// A timestamped backup of the original file is created first, and the
        /// workbook is saved again in its original file format.
        /// </summary>
        public static void Cleanup()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "WFDOT Excel-Datei zum Bereinigen auswählen";
                dialog.Filter = "Excel-Arbeitsmappe (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|Alle Dateien (*.*)|*.*";
                Log("Datei-Dialog geöffnet, warte auf Auswahl ...");
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    Log("Cleanup abgebrochen (keine Datei ausgewählt).");
                    return;
                }

                Log($"Ausgewählte Datei: {dialog.FileName}");
                Cleanup(dialog.FileName);
            }
        }

        public static void Cleanup(string excelPath)
        {
            var stopwatch = Stopwatch.StartNew();
            Log("=== Excel-Cleanup gestartet ===");
            Log($"Excel-Datei : {excelPath}");

            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("Excel-Datei nicht gefunden.", excelPath);
            }

            // Back up the original before overwriting it.
            var backupPath = Path.Combine(
                Path.GetDirectoryName(excelPath),
                Path.GetFileNameWithoutExtension(excelPath) + "_backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(excelPath));
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

                Log("Lese den gesamten Bereich auf einmal aus Excel ...");
                var readWatch = Stopwatch.StartNew();
                // Values are used to compute the changes.
                var cells = (object[,])xlRange.Value2;
                // Formulas are used to detect formula cells, which must never be touched.
                // (Value2 only returns the computed result, so a full-range write would
                //  silently replace every formula with its cached value.)
                var formulas = (object[,])xlRange.Formula;
                Log($"Bereich ausgelesen in {readWatch.Elapsed.TotalSeconds:F1} s.");

                // A cell is a formula cell when its formula differs from its value, i.e. it
                // starts with "=". Such cells keep their formula and are never modified.
                Func<int, int, bool> isFormula = (row, col) =>
                {
                    if (formulas[row, col] == null)
                    {
                        return false;
                    }

                    return Convert.ToString(formulas[row, col]).StartsWith("=", StringComparison.Ordinal);
                };

                // Find the Ostfriesisch column so the clitic rule is applied only there.
                var headerToColumn = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    string header = CellText(cells, 1, c);
                    if (!string.IsNullOrWhiteSpace(header) && !headerToColumn.ContainsKey(header))
                    {
                        headerToColumn[header] = c;
                    }
                }

                // Header -> column name in the database (used to decide the clitic rule).
                var columnNames = new string[colCount + 1];
                foreach (var pair in headerToColumn)
                {
                    columnNames[pair.Value] = pair.Key;
                }

                Log("Bereinige Zellen (Leerzeichen am Ende, Ostfriesisch-Klitika) ...");

                // Phase 1: compute the cleaned value for every cell and collect the changes.
                // Nothing is written yet. Formula cells are skipped so their formulas survive.
                var changes = new List<CellChange>();
                int skippedFormulaCells = 0;
                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        if (isFormula(r, c))
                        {
                            skippedFormulaCells++;
                            continue;
                        }

                        string column = columnNames[c];
                        string original = CellText(cells, r, c);

                        bool isOstfriesisch = string.Equals(column, "Ostfriesisch", StringComparison.OrdinalIgnoreCase);
                        string value = CleanForExcel(isOstfriesisch, original, out bool apostropheTouched, out bool whitespaceTrimmed);

                        if (apostropheTouched || whitespaceTrimmed)
                        {
                            changes.Add(new CellChange
                            {
                                Row = r,
                                Column = c,
                                ColumnName = column,
                                Original = original,
                                ExcelValue = value,
                                // Excel drops a single leading apostrophe (it becomes the "text"
                                // marker), so the value actually stored is value without it.
                                DisplayValue = value.StartsWith("''", StringComparison.Ordinal) ? value.Substring(1) : value,
                                IsApostrophe = apostropheTouched
                            });
                        }
                    }

                    if (r % 10000 == 0)
                    {
                        Log($"  ... {r - 1} Zeilen geprüft");
                    }
                }

                if (skippedFormulaCells > 0)
                {
                    Log($"Formelzellen übersprungen (bleiben unverändert): {skippedFormulaCells}");
                }

                int apostrophes = 0;
                int whitespaceOnly = 0;
                foreach (var change in changes)
                {
                    if (change.IsApostrophe)
                    {
                        apostrophes++;
                    }
                    else
                    {
                        whitespaceOnly++;
                    }
                }

                Log($"Gefundene Änderungen: {changes.Count}" +
                    (apostrophes > 0 ? $", davon Apostroph-Korrekturen: {apostrophes}" : "") +
                    (whitespaceOnly > 0 ? $", davon nur Leerzeichen am Ende: {whitespaceOnly}" : ""));

                // Ask once whether trailing whitespace may be cleaned without per-cell confirmation.
                bool autoWhitespace = false;
                if (whitespaceOnly > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine($"Es wurden {whitespaceOnly} Zelle(n) mit Leerzeichen am Ende gefunden.");
                    Console.Write("Sollen diese Leerzeichen automatisch (ohne Einzelabfrage) entfernt werden? (j/n): ");
                    string answer = (Console.ReadLine() ?? string.Empty).Trim().ToLowerInvariant();
                    autoWhitespace = answer == "j" || answer == "ja" || answer == "y";
                    Log(autoWhitespace
                        ? "Leerzeichen am Ende werden automatisch entfernt."
                        : "Leerzeichen am Ende werden einzeln abgefragt.");
                }

                // Phase 2: let the user confirm every single change.
                int approved = 0;
                int rejected = 0;
                bool approveAll = false;
                bool cancelled = false;

                // Whitespace-only changes that still need asking (when auto is off).
                var toAsk = new List<CellChange>();
                foreach (var change in changes)
                {
                    if (!change.IsApostrophe && autoWhitespace)
                    {
                        change.Approved = true;
                        approved++;
                        continue;
                    }

                    toAsk.Add(change);
                }

                if (toAsk.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine($"Es gibt {toAsk.Count} zu ändernde Zelle(n), die bestätigt werden müssen.");
                    Console.WriteLine("Antworten: j = übernehmen, n = unverändert lassen, a = alle restlichen übernehmen, q = abbrechen (nichts speichern)");
                    Console.WriteLine();

                    int index = 0;
                    foreach (var change in toAsk)
                    {
                        index++;
                        Console.WriteLine($"--- Änderung {index} von {toAsk.Count} ---");
                        Console.WriteLine($"  Zeile {change.Row}, Spalte {change.Column} (\"{change.ColumnName}\")");
                        Console.WriteLine($"  alt:   \"{change.Original}\"");
                        Console.WriteLine($"  neu:   \"{change.DisplayValue}\"" + (change.IsApostrophe ? "   [Apostroph-Korrektur]" : "   [Leerzeichen am Ende]"));

                        if (approveAll)
                        {
                            change.Approved = true;
                            approved++;
                            Console.WriteLine("  -> übernommen (a)");
                            continue;
                        }

                        Console.Write("  übernehmen? (j/n/a/q): ");
                        string answer = (Console.ReadLine() ?? string.Empty).Trim().ToLowerInvariant();
                        Console.WriteLine();

                        switch (answer)
                        {
                            case "j":
                            case "ja":
                            case "y":
                                change.Approved = true;
                                approved++;
                                break;
                            case "a":
                            case "alle":
                                approveAll = true;
                                change.Approved = true;
                                approved++;
                                break;
                            case "q":
                            case "quit":
                            case "abbrechen":
                                cancelled = true;
                                Log("Abbruch durch Benutzer.");
                                break;
                            default:
                                change.Approved = false;
                                rejected++;
                                break;
                        }

                        if (cancelled)
                        {
                            break;
                        }
                    }
                }

                if (!cancelled)
                {
                    Log($"Übernommen: {approved}, abgelehnt: {rejected}" +
                        (changes.Count == 0 ? " (keine Änderungen nötig)" : ""));
                }

                if (approved > 0)
                {
                    // Phase 3: write back ONLY the approved cells. Writing the whole range back
                    // would overwrite formulas and every untouched cell, so the changes are
                    // grouped per column into contiguous blocks and each block is written with
                    // a single Range call (fast and reliable, unlike hundreds of single-cell writes).
                    Log("Schreibe bestätigte Werte zurück in die Excel-Datei (nur die geänderten Zellen) ...");
                    int writtenBlocks = 0;
                    int writtenCells = 0;

                    var byColumn = new Dictionary<int, List<CellChange>>();
                    foreach (var change in changes)
                    {
                        if (!change.Approved)
                        {
                            continue;
                        }

                        if (!byColumn.TryGetValue(change.Column, out var list))
                        {
                            list = new List<CellChange>();
                            byColumn[change.Column] = list;
                        }
                        list.Add(change);
                    }

                    foreach (var columnPair in byColumn)
                    {
                        int column = columnPair.Key;
                        var columnChanges = columnPair.Value;
                        columnChanges.Sort((a, b) => a.Row.CompareTo(b.Row));

                        // Group adjacent rows into contiguous blocks.
                        int i = 0;
                        while (i < columnChanges.Count)
                        {
                            int j = i;
                            while (j + 1 < columnChanges.Count && columnChanges[j + 1].Row == columnChanges[j].Row + 1)
                            {
                                j++;
                            }

                            int firstRow = columnChanges[i].Row;
                            int lastRow = columnChanges[j].Row;
                            int blockSize = lastRow - firstRow + 1;

                            // One row per block row, one column: a (blockSize x 1) array, 0-based.
                            var blockValues = new object[blockSize, 1];
                            for (int k = 0; k < blockSize; k++)
                            {
                                blockValues[k, 0] = columnChanges[i + k].ExcelValue;
                            }

                            Excel.Range block = null;
                            try
                            {
                                block = (Excel.Range)xlWorksheet.Range[
                                    xlWorksheet.Cells[firstRow, column],
                                    xlWorksheet.Cells[lastRow, column]];
                                block.Value2 = blockValues;
                                writtenBlocks++;
                                writtenCells += blockSize;
                            }
                            finally
                            {
                                if (block != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(block);
                                }
                            }

                            i = j + 1;
                        }
                    }
                    Log($"  Geschriebene Zellen: {writtenCells} (in {writtenBlocks} Block(en))");

                    Log("Speichere Excel-Datei ...");
                    xlWorkbook.Save();
                    Log("Excel-Datei gespeichert.");

                    // Verify AFTER saving: re-open the workbook and read the values back.
                    // Reading from the still-open instance can return stale cached values.
                    Log("Verifiziere gespeicherte Datei (erneut öffnen) ...");
                    VerifySavedCells(excelPath, changes);
                }
                else if (cancelled)
                {
                    Log("Abbruch: Excel-Datei bleibt unverändert (nichts gespeichert).");
                }
                else if (changes.Count == 0)
                {
                    Log("Keine Änderungen nötig; Excel-Datei bleibt unverändert.");
                }
                else
                {
                    Log("Keine Änderung bestätigt; Excel-Datei bleibt unverändert.");
                }
            }
            catch (Exception ex)
            {
                Log($"=== FEHLER: {ex.Message} ===");
                Log($"Die Originaldatei wurde nicht verändert. Backup siehe oben.");
                throw;
            }
            finally
            {
                Log("Schließe Excel und räume auf ...");
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                }
                if (xlWorksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                }
                if (xlRange != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Log("Aufgeräumt.");
            }

            stopwatch.Stop();
            Log($"=== Excel-Cleanup erfolgreich abgeschlossen in {stopwatch.Elapsed.TotalSeconds:F1} s ===");
        }

        /// <summary>Reads a cell from the prefetched range array as text.</summary>
        private static string CellText(object[,] cells, int row, int column)
        {
            object value = cells[row, column];
            return value == null ? string.Empty : Convert.ToString(value);
        }

        /// <summary>
        /// Re-opens the saved workbook and checks that the approved cells really hold the
        /// intended value. Verifying in the still-open Excel instance can return stale
        /// values, so the file is read back from disk.
        /// </summary>
        private static void VerifySavedCells(string excelPath, List<CellChange> changes)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                // Read-only so the verification can never modify anything.
                xlWorkbook = xlApp.Workbooks.Open(excelPath, ReadOnly: true);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;
                var cells = (object[,])xlRange.Value2;

                int checkedCells = 0;
                int mismatches = 0;
                foreach (var change in changes)
                {
                    if (!change.Approved)
                    {
                        continue;
                    }

                    checkedCells++;
                    string actual = CellText(cells, change.Row, change.Column);
                    if (actual != change.DisplayValue)
                    {
                        mismatches++;
                        if (mismatches <= 10)
                        {
                            Log($"  WARNUNG: Zeile {change.Row}, Spalte {change.Column} (\"{change.ColumnName}\") nicht gespeichert: erwartet \"{change.DisplayValue}\", gelesen \"{actual}\"");
                        }
                    }
                }

                if (mismatches > 0)
                {
                    Log($"  ACHTUNG: {mismatches} von {checkedCells} Zellen weichen nach dem Speichern ab.");
                }
                else
                {
                    Log($"  Speichern verifiziert ({checkedCells} Zellen geprüft).");
                }
            }
            finally
            {
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                }
                if (xlWorksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                }
                if (xlRange != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// One proposed cell change found by the Excel cleanup. The user confirms it
        /// before it is written back into the workbook.
        /// </summary>
        private sealed class CellChange
        {
            /// <summary>Worksheet row (1-based, including the header row).</summary>
            public int Row { get; set; }

            /// <summary>Worksheet column (1-based).</summary>
            public int Column { get; set; }

            /// <summary>Header name of the column ("Ostfriesisch", "Deutsch", ...).</summary>
            public string ColumnName { get; set; }

            /// <summary>Value currently in the cell.</summary>
            public string Original { get; set; }

            /// <summary>Raw string to write into the cell (apostrophe may be doubled).</summary>
            public string ExcelValue { get; set; }

            /// <summary>Value the user sees in Excel after writing.</summary>
            public string DisplayValue { get; set; }

            /// <summary>True when this is a real clitic apostrophe correction.</summary>
            public bool IsApostrophe { get; set; }

            /// <summary>True when the user confirmed the change.</summary>
            public bool Approved { get; set; }
        }

        /// <summary>
        /// Applies the cleanup rules to a single cell value. The leading-apostrophe
        /// (clitic) rule is only applied to the Ostfriesisch column, matching the
        /// existing database content.
        /// </summary>
        /// <param name="column">Name of the database column.</param>
        /// <param name="value">Raw cell value.</param>
        /// <param name="apostropheAdded">True when a leading apostrophe was inserted (a real correction).</param>
        /// <param name="whitespaceTrimmed">True when trailing whitespace was removed (mere normalisation).</param>
        private static string Clean(string column, string value, out bool apostropheAdded, out bool whitespaceTrimmed)
        {
            apostropheAdded = false;
            whitespaceTrimmed = false;

            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            // Trailing whitespace is not meaningful; internal content is kept as-is.
            string trimmed = value.TrimEnd('\r', '\n', ' ');
            if (trimmed.Length != value.Length)
            {
                whitespaceTrimmed = true;
                value = trimmed;
            }

            if (column == "Ostfriesisch")
            {
                // A doubled leading apostrophe is how a real apostrophe survives in Excel
                // (see CleanForExcel). Unpack it back to a single apostrophe for the database.
                if (value.StartsWith("''", StringComparison.Ordinal))
                {
                    value = value.Substring(1);
                }

                if (CliticRule.IsMatch(value) && !value.StartsWith("'", StringComparison.Ordinal))
                {
                    apostropheAdded = true;
                    return "'" + value;
                }
            }

            return value;
        }

        /// <summary>
        /// Cleanup rule for the Excel file itself. Identical to <see cref="Clean"/> except that
        /// a leading apostrophe is written doubled ("''k"), because Excel treats a single
        /// leading apostrophe as its "this is text" marker and drops it when saving.
        /// Excel then stores and displays "'k", and the import collapses it back to "'k".
        /// </summary>
        /// <param name="isOstfriesisch">True when the cell belongs to the Ostfriesisch column.</param>
        /// <param name="value">Raw cell value.</param>
        /// <param name="apostropheTouched">True when a leading apostrophe was inserted or doubled.</param>
        /// <param name="whitespaceTrimmed">True when trailing whitespace was removed.</param>
        private static string CleanForExcel(bool isOstfriesisch, string value, out bool apostropheTouched, out bool whitespaceTrimmed)
        {
            apostropheTouched = false;
            whitespaceTrimmed = false;

            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            string trimmed = value.TrimEnd('\r', '\n', ' ');
            if (trimmed.Length != value.Length)
            {
                whitespaceTrimmed = true;
                value = trimmed;
            }

            if (isOstfriesisch)
            {
                // Already has an apostrophe ('k or ''k) -> leave it completely alone.
                // Doubling it here would only produce a change that looks like a no-op
                // (alt == neu) and would show up again on every run.
                if (value.StartsWith("'", StringComparison.Ordinal))
                {
                    return value;
                }

                if (CliticRule.IsMatch(value))
                {
                    apostropheTouched = true;
                    return "''" + value;
                }
            }

            return value;
        }

        /// <summary>Drops and recreates the FTS4 index and refills it from WB.</summary>
        private static void RebuildFts(SQLiteConnection connection)
        {
            Log("  Lösche alten Suchindex (WBFTS) ...");
            Exec(connection, null, "DROP TABLE IF EXISTS WBFTS;");
            Log("  Erstelle Suchindex (WBFTS) neu ...");
            Exec(connection, null,
                "CREATE VIRTUAL TABLE WBFTS USING fts4(ID, Ostfriesisch, Deutsch, Artikel, Wortart, " +
                "Plural, Genus, Komparation, Konjugation, Nebenformen, Standardform, Zuordnung, Nummer, " +
                "Englisch, tokenize=unicode61);");
            Log("  Fülle Suchindex aus Tabelle WB ...");
            Exec(connection, null,
                "INSERT INTO WBFTS (" + string.Join(", ", FtsColumns) + ") SELECT " +
                string.Join(", ", FtsColumns) + " FROM WB;");
            Log("  Optimiere Suchindex ...");
            Exec(connection, null, "INSERT INTO WBFTS(WBFTS) VALUES ('optimize');");
        }

        private static void Exec(SQLiteConnection connection, SQLiteTransaction transaction, string sql)
        {
            using (var command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = sql;
                command.ExecuteNonQuery();
            }
        }
    }
}
