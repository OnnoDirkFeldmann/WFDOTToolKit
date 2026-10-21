using System;
using System.Collections.Generic;
using System.Data.SQLite;
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

        public static void Import()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "WFDOT Excel-Datei auswählen";
                dialog.Filter = "Excel-Arbeitsmappe (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|Alle Dateien (*.*)|*.*";
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    Console.WriteLine("Import abgebrochen.");
                    return;
                }

                Import(dialog.FileName, WfdotPath);
            }
        }

        /// <summary>Absolute path of the database that lives next to the executable's source folder.</summary>
        private static string WfdotPath =>
            @"D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";

        public static void Import(string excelPath, string sqlitePath)
        {
            // Operate on a copy so the original workbook stays untouched even if Excel offers to save.
            var workbookCopy = Path.Combine(Path.GetTempPath(), "wfdot_import_" + Guid.NewGuid().ToString("N") + Path.GetExtension(excelPath));
            File.Copy(excelPath, workbookCopy, true);

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                xlWorkbook = xlApp.Workbooks.Open(workbookCopy, ReadOnly: true);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                // Map header names -> worksheet column index, honouring the ignore list.
                var headerToColumn = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    string header = CellText(xlWorksheet.Cells[1, c]);
                    if (string.IsNullOrWhiteSpace(header) || IgnoredColumns.Contains(header))
                    {
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

                Console.WriteLine($"Lese {rowCount - 1} Zeilen aus {Path.GetFileName(excelPath)} ...");

                using (var connection = new SQLiteConnection("Data Source=" + sqlitePath))
                {
                    connection.Open();

                    using (var transaction = connection.BeginTransaction())
                    {
                        Exec(connection, transaction,
                            "CREATE TABLE IF NOT EXISTS WB (ID INTEGER NOT NULL UNIQUE, " +
                            "Ostfriesisch TEXT NOT NULL, Deutsch TEXT NOT NULL, Artikel TEXT NOT NULL, " +
                            "Wortart TEXT NOT NULL, Plural TEXT NOT NULL, Genus TEXT NOT NULL, " +
                            "Komparation TEXT NOT NULL, Konjugation TEXT NOT NULL, Nebenformen TEXT NOT NULL, " +
                            "Standardform TEXT NOT NULL, Zuordnung TEXT NOT NULL, Nummer TEXT NOT NULL, " +
                            "Etymologie TEXT NOT NULL, Kommentar TEXT NOT NULL, Rezept TEXT NOT NULL, " +
                            "Englisch TEXT NOT NULL, PRIMARY KEY (ID));");
                        Exec(connection, transaction, "DELETE FROM WB;");

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
                                    string value = Clean(DbColumns[i], CellText(xlWorksheet.Cells[r, headerToColumn[DbColumns[i]]]));
                                    parameters[i].Value = value;
                                    if (value.Length > 0)
                                    {
                                        empty = false;
                                    }
                                }

                                if (empty)
                                {
                                    continue; // skip fully empty rows
                                }

                                insert.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }

                    RebuildFts(connection);
                    Console.WriteLine("Import abgeschlossen.");
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

                try { File.Delete(workbookCopy); } catch { /* best effort */ }
            }
        }

        /// <summary>Reads a cell as text, converting the COM value safely.</summary>
        private static string CellText(Excel.Range cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }
            object value = cell.Value2;
            return value == null ? string.Empty : Convert.ToString(value);
        }

        /// <summary>
        /// Applies the cleanup rules to a single cell value. The leading-apostrophe
        /// (clitic) rule is only applied to the Ostfriesisch column, matching the
        /// existing database content.
        /// </summary>
        private static string Clean(string column, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            // Trailing whitespace is not meaningful; internal content is kept as-is.
            value = value.TrimEnd('\r', '\n', ' ');

            if (column == "Ostfriesisch" && CliticRule.IsMatch(value))
            {
                return "'" + value;
            }

            return value;
        }

        /// <summary>Drops and recreates the FTS4 index and refills it from WB.</summary>
        private static void RebuildFts(SQLiteConnection connection)
        {
            Exec(connection, null, "DROP TABLE IF EXISTS WBFTS;");
            Exec(connection, null,
                "CREATE VIRTUAL TABLE WBFTS USING fts4(ID, Ostfriesisch, Deutsch, Artikel, Wortart, " +
                "Plural, Genus, Komparation, Konjugation, Nebenformen, Standardform, Zuordnung, Nummer, " +
                "Englisch, tokenize=unicode61);");
            Exec(connection, null,
                "INSERT INTO WBFTS (" + string.Join(", ", FtsColumns) + ") SELECT " +
                string.Join(", ", FtsColumns) + " FROM WB;");
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
