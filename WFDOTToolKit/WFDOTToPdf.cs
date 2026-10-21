using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace WFDOTToolKit
{
    /// <summary>
    /// Exports a generated docx file to PDF with the Word installation of the
    /// machine. Word is addressed through late binding, so the project does not
    /// need a build time reference to the Office interop assemblies and the
    /// export also works with different Word versions.
    /// </summary>
    internal static class WFDOTToPdf
    {
        /// <summary>wdExportFormatPDF</summary>
        private const int ExportFormatPdf = 17;

        /// <summary>wdDoNotSaveChanges</summary>
        private const int DoNotSaveChanges = 0;

        /// <summary>wdAlertsNone</summary>
        private const int AlertsNone = 0;

        /// <summary>
        /// Converts <paramref name="docxPath"/> into a PDF file at
        /// <paramref name="pdfPath"/>.
        /// </summary>
        public static void ExportToPdf(string docxPath, string pdfPath)
        {
            if (!File.Exists(docxPath))
            {
                throw new FileNotFoundException("DOCX-Datei nicht gefunden.", docxPath);
            }

            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                throw new InvalidOperationException(
                    "Microsoft Word ist auf diesem Rechner nicht installiert. " +
                    "Ohne Word kann keine PDF-Datei erzeugt werden.");
            }

            var watch = Stopwatch.StartNew();
            Log($"Exportiere PDF über Microsoft Word ({Path.GetFileName(docxPath)}) ...");

            object wordObject = null;
            object documentObject = null;
            dynamic word = null;
            dynamic document = null;

            try
            {
                wordObject = Activator.CreateInstance(wordType);
                word = wordObject;
                word.Visible = false;
                word.DisplayAlerts = AlertsNone;

                // FileName, ConfirmConversions, ReadOnly
                documentObject = word.Documents.Open(Path.GetFullPath(docxPath), false, true);
                document = documentObject;

                document.ExportAsFixedFormat(Path.GetFullPath(pdfPath), ExportFormatPdf);

                document.Close(DoNotSaveChanges);
                documentObject = null;
                document = null;

                word.Quit(DoNotSaveChanges);
                wordObject = null;
                word = null;
            }
            finally
            {
                try
                {
                    if (documentObject != null)
                    {
                        ((dynamic)documentObject).Close(DoNotSaveChanges);
                    }
                }
                catch
                {
                    // best effort
                }

                try
                {
                    if (wordObject != null)
                    {
                        ((dynamic)wordObject).Quit(DoNotSaveChanges);
                    }
                }
                catch
                {
                    // best effort
                }

                if (documentObject != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(documentObject);
                    }
                    catch
                    {
                        // best effort
                    }
                }

                if (wordObject != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(wordObject);
                    }
                    catch
                    {
                        // best effort
                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            var size = new FileInfo(pdfPath).Length / 1024.0 / 1024.0;
            Log($"PDF erzeugt: {pdfPath} ({size:F1} MB, {watch.Elapsed.TotalSeconds:F1} s).");
        }

        private static void Log(string message)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }
    }
}
