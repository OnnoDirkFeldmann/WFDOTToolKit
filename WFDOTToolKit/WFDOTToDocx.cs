using Novacode;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WFDOTToolKit
{
    /// <summary>
    /// Builds the Word output of the dictionary.
    ///
    /// Output layout:
    ///   out0.docx, out1.docx, ...  one file per <see cref="EntriesPerFile"/> entries
    ///   outOhneZu.docx             phrases without an assignment
    ///   full.docx                  all part files merged together
    ///
    /// Every file uses a two-column layout and shows the page number in the
    /// bottom right corner.
    ///
    /// To keep this fast with 100k+ entries:
    ///   * The database is read once into lightweight objects (no DataSet).
    ///   * Phrases are looked up through a dictionary (O(1) instead of O(n) per entry).
    ///   * Paragraphs are built as plain WordprocessingML. The InsertText API of
    ///     the document library gets slower with every run in the document, which
    ///     is what made the old version take minutes.
    ///   * The part files are generated in parallel.
    ///   * The full document is assembled on XML level instead of loading and
    ///     copying complete documents again.
    /// </summary>
    public class WFDOTToDocx
    {
        /// <summary>Path of the dictionary database.</summary>
        private const string DatabasePath = @"D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";

        /// <summary>Entries per part file (out0.docx, out1.docx, ...).</summary>
        private const int EntriesPerFile = 1000;

        /// <summary>Name of the merged document (the actual output file).</summary>
        public const string MainFileName = "WFDOT.docx";

        /// <summary>Page margin top and bottom in twips (2 cm).</summary>
        private const int MarginTopBottom = 1134;

        /// <summary>Page margin left and right in twips (2.54 cm).</summary>
        private const int MarginLeftRight = 1440;

        /// <summary>Font of the headword and the entry text.</summary>
        private const string FontFamily = "Segoe UI";

        private const double EntryFontSize = 9D;
        private const double PhraseFontSize = 8D;

        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static readonly XNamespace W = WordNamespace;

        /// <summary>
        /// Word classes that are not shown in square brackets in the dictionary.
        /// Every other word class produces the text "Wortartfehler".
        /// </summary>
        private static readonly HashSet<string> WordClassesWithoutLabel = new HashSet<string>(StringComparer.Ordinal)
        {
            "Abkürzung", "Adjektiv", "Adverb", "Artikel", "Ausruf", "Flexionsform",
            "Interrogativpronomen", "Konjunktion", "Nachsilbe", "Name", "Numeral",
            "Ortsname", "Partikel", "Partizip", "Phrase", "Pronomen",
            "Pronominaladverb", "Substantiv", "Verb", "Vorsilbe", "Zwischensilbe"
        };

        public static void ToDocx(bool fullversion, string path, bool testMode)
        {
            var totalWatch = Stopwatch.StartNew();

            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
            Directory.CreateDirectory(path);

            Log("Lese Datenbank ...");
            var loadWatch = Stopwatch.StartNew();
            List<Entry> entries;
            Dictionary<string, List<Entry>> phraseIndex;
            List<Entry> unassignedPhrases;
            LoadData(testMode, out entries, out phraseIndex, out unassignedPhrases);
            Log($"  {entries.Count} Einträge und {phraseIndex.Count} Phrasen-Zuordnungen in " +
                $"{loadWatch.Elapsed.TotalSeconds:F1} s gelesen.");

            var files = new List<string>();
            var fileCount = Math.Max(1, (entries.Count + EntriesPerFile - 1) / EntriesPerFile);
            var fileNames = new string[fileCount];
            for (var index = 0; index < fileCount; index++)
            {
                fileNames[index] = Path.Combine(path, $"out{index}.docx");
            }

            Log($"Erzeuge {fileCount} Teildateien ...");
            var generationWatch = Stopwatch.StartNew();
            var done = 0;
            Parallel.For(0, fileCount, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, index =>
            {
                var from = index * EntriesPerFile;
                var count = Math.Min(EntriesPerFile, entries.Count - from);
                GenerateFile(fileNames[index], entries, from, count, fullversion, phraseIndex);
                var finished = System.Threading.Interlocked.Increment(ref done);
                if (finished % 10 == 0 || finished == fileCount)
                {
                    Log($"  ... {finished} von {fileCount} Teildateien");
                }
            });
            files.AddRange(fileNames);
            Log($"  Teildateien erzeugt in {generationWatch.Elapsed.TotalSeconds:F1} s.");

            if (unassignedPhrases.Count > 0)
            {
                Log($"Erzeuge Datei der unzugeordneten Phrasen ({unassignedPhrases.Count}) ...");
                var fileUnassigned = Path.Combine(path, "outOhneZu.docx");
                WriteUnassignedPhrases(fileUnassigned, unassignedPhrases);
                files.Add(fileUnassigned);
            }

            Log("Füge Teildateien zu " + MainFileName + " zusammen ...");
            var mergeWatch = Stopwatch.StartNew();
            MergeFiles(GetMainFile(path), files);
            Log($"  {MainFileName} erzeugt in {mergeWatch.Elapsed.TotalSeconds:F1} s.");

            Log($"Fertig: {files.Count + 1} Dateien in {path} " +
                $"(Gesamtzeit {totalWatch.Elapsed.TotalSeconds:F1} s).");
        }

        /// <summary>Returns the path of the merged document inside the output folder.</summary>
        public static string GetMainFile(string path)
        {
            return Path.Combine(path, MainFileName);
        }

        /// <summary>A dictionary entry with only the fields the output needs.</summary>
        private sealed class Entry
        {
            public string Ostfriesisch;
            public string Deutsch;
            public string Artikel;
            public string Wortart;
            public string Plural;
            public string Genus;
            public string Komparation;
            public string Konjugation;
            public string Nebenformen;
            public string Standardform;
            public string Zuordnung;
            public string Nummer;
            public string Kommentar;
            public string Rezept;
        }

        /// <summary>
        /// Reads all entries from the database and pre-sorts them into regular
        /// entries, phrases per assignment, and phrases without an assignment.
        /// </summary>
        private static void LoadData(
            bool testMode,
            out List<Entry> entries,
            out Dictionary<string, List<Entry>> phraseIndex,
            out List<Entry> unassignedPhrases)
        {
            entries = new List<Entry>(100000);
            phraseIndex = new Dictionary<string, List<Entry>>(40000, StringComparer.Ordinal);
            unassignedPhrases = new List<Entry>(4000);

            using (var connection = new SQLiteConnection("Data Source=" + DatabasePath))
            {
                connection.Open();
                var sql = testMode ? "SELECT * FROM WB WHERE Ostfriesisch Like 'a%' or Ostfriesisch Like 'â%'" : "SELECT * FROM WB";
                using (var command = new SQLiteCommand(sql, connection))
                using (var reader = command.ExecuteReader())
                {
                    var ordinalOstfriesisch = reader.GetOrdinal("Ostfriesisch");
                    var ordinalDeutsch = reader.GetOrdinal("Deutsch");
                    var ordinalArtikel = reader.GetOrdinal("Artikel");
                    var ordinalWortart = reader.GetOrdinal("Wortart");
                    var ordinalPlural = reader.GetOrdinal("Plural");
                    var ordinalGenus = reader.GetOrdinal("Genus");
                    var ordinalKomparation = reader.GetOrdinal("Komparation");
                    var ordinalKonjugation = reader.GetOrdinal("Konjugation");
                    var ordinalNebenformen = reader.GetOrdinal("Nebenformen");
                    var ordinalStandardform = reader.GetOrdinal("Standardform");
                    var ordinalZuordnung = reader.GetOrdinal("Zuordnung");
                    var ordinalNummer = reader.GetOrdinal("Nummer");
                    var ordinalKommentar = reader.GetOrdinal("Kommentar");
                    var ordinalRezept = reader.GetOrdinal("Rezept");

                    while (reader.Read())
                    {
                        var entry = new Entry
                        {
                            Ostfriesisch = ReadString(reader, ordinalOstfriesisch),
                            Deutsch = ReadString(reader, ordinalDeutsch),
                            Artikel = ReadString(reader, ordinalArtikel),
                            Wortart = ReadString(reader, ordinalWortart),
                            Plural = ReadString(reader, ordinalPlural),
                            Genus = ReadString(reader, ordinalGenus),
                            Komparation = ReadString(reader, ordinalKomparation),
                            Konjugation = ReadString(reader, ordinalKonjugation),
                            Nebenformen = ReadString(reader, ordinalNebenformen),
                            Standardform = ReadString(reader, ordinalStandardform),
                            Zuordnung = ReadString(reader, ordinalZuordnung),
                            Nummer = ReadString(reader, ordinalNummer),
                            Kommentar = ReadString(reader, ordinalKommentar),
                            Rezept = ReadString(reader, ordinalRezept)
                        };

                        if (entry.Wortart == "Phrase")
                        {
                            if (entry.Zuordnung == "-")
                            {
                                unassignedPhrases.Add(entry);
                            }
                            else
                            {
                                if (!phraseIndex.TryGetValue(entry.Zuordnung, out var list))
                                {
                                    list = new List<Entry>(1);
                                    phraseIndex[entry.Zuordnung] = list;
                                }
                                list.Add(entry);
                            }
                        }
                        else
                        {
                            entries.Add(entry);
                        }
                    }
                }
            }
        }

        /// <summary>Reads a text column; a missing value counts as "-" like an empty cell.</summary>
        private static string ReadString(SQLiteDataReader reader, int ordinal)
        {
            return reader.IsDBNull(ordinal) ? "-" : reader.GetString(ordinal);
        }

        /// <summary>Creates one part file with <paramref name="count"/> entries starting at <paramref name="from"/>.</summary>
        private static void GenerateFile(
            string file,
            List<Entry> entries,
            int from,
            int count,
            bool fullversion,
            Dictionary<string, List<Entry>> phraseIndex)
        {
            var doc = DocX.Create(file);
            var body = doc.Xml;
            var noteFormat = CreateFormat(PhraseFontSize, false, false);

            for (var i = from; i < from + count; i++)
            {
                WriteEntry(doc, body, entries[i], fullversion, phraseIndex, noteFormat);
            }

            ApplyLayout(doc);
            doc.Save();
            doc.Dispose();
        }

        private static Formatting CreateFormat(double size, bool bold, bool italic)
        {
            var formatting = new Formatting();
            formatting.FontFamily = new Font(FontFamily);
            formatting.Size = size;
            formatting.Bold = bold;
            formatting.Italic = italic;
            return formatting;
        }

        /// <summary>
        /// Writes one entry as a paragraph. The headword is bold and everything else
        /// goes into a single run.
        /// </summary>
        private static void WriteEntry(
            DocX doc,
            XElement body,
            Entry row,
            bool fullversion,
            Dictionary<string, List<Entry>> phraseIndex,
            Formatting noteFormat)
        {
            var text = new StringBuilder(96);
            var wordClass = row.Wortart != "-" && !WordClassesWithoutLabel.Contains(row.Wortart)
                ? "Wortartfehler"
                : "-";

            var hasArticle = row.Artikel != "-";
            var hasGenus = row.Genus != "-";
            var hasPlural = row.Plural != "-";
            var hasWordClass = wordClass != "-";
            var hasBracket = hasArticle || hasGenus || hasPlural || (row.Wortart != "-" && hasWordClass);

            if (hasBracket)
            {
                text.Append(" [");
            }
            if (hasArticle)
            {
                text.Append(row.Artikel);
            }
            if (hasWordClass)
            {
                if (hasArticle)
                {
                    text.Append(", ");
                }
                text.Append(wordClass);
            }
            if (hasGenus)
            {
                if (hasArticle || hasWordClass)
                {
                    text.Append(", ");
                }
                text.Append(GenusText(row.Genus));
            }
            if (hasPlural)
            {
                text.Append(", ");
                text.Append(row.Plural);
            }
            if (hasBracket)
            {
                text.Append(']');
            }

            text.Append(' ');
            text.Append(row.Deutsch);

            if (row.Nebenformen != "-")
            {
                text.Append(Environment.NewLine);
                text.Append("[NF: ").Append(row.Nebenformen).Append(']');
            }
            if (row.Standardform != "-")
            {
                text.Append(Environment.NewLine);
                text.Append('[').Append(row.Standardform).Append(']');
            }

            var paragraph = CreateParagraph();
            paragraph.Add(CreateRun(row.Ostfriesisch, EntryFontSize, true, false));
            paragraph.Add(CreateRun(text.ToString(), EntryFontSize, false, false));
            body.Add(paragraph);

            if (fullversion)
            {
                if (row.Komparation != "-")
                {
                    InsertNoteTable(doc, ShortenHtml(row.Komparation, 5).Replace("<br/>", "; "), noteFormat);
                }
                if (row.Konjugation != "-")
                {
                    InsertNoteTable(doc, ShortenHtml(row.Konjugation, 5).Replace("<br/>", "; "), noteFormat);
                }
                if (row.Kommentar != "-")
                {
                    InsertNoteTable(doc, row.Kommentar, noteFormat);
                }
                if (row.Rezept != "-")
                {
                    InsertNoteTable(doc, row.Rezept.Replace("<br/>", " "), noteFormat);
                }
            }

            // Attach the phrases that belong to this entry.
            var phraseKey = row.Nummer == "-" ? row.Ostfriesisch : row.Ostfriesisch + "=" + row.Nummer;
            if (phraseIndex.TryGetValue(phraseKey, out var phrases))
            {
                var phraseParagraph = CreateParagraph();
                for (var j = 0; j < phrases.Count; j++)
                {
                    var phrase = phrases[j];
                    phraseParagraph.Add(CreateRun(phrase.Ostfriesisch, PhraseFontSize, true, false));
                    phraseParagraph.Add(CreateRun(" ", PhraseFontSize, false, false));
                    phraseParagraph.Add(CreateRun(phrase.Deutsch, PhraseFontSize, false, true));
                    if (j != phrases.Count - 1)
                    {
                        phraseParagraph.Add(CreateRun(" ", PhraseFontSize, false, false));
                    }
                }
                body.Add(phraseParagraph);
            }
        }

        /// <summary>Removes the trailing HTML remainder ("&lt;br/&gt;") of a note.</summary>
        private static string ShortenHtml(string value, int trimEnd)
        {
            return value.Length > trimEnd ? value.Substring(0, value.Length - trimEnd) : value;
        }

        private static string GenusText(string genus)
        {
            switch (genus)
            {
                case "m":
                    return "m.";
                case "f":
                    return "f.";
                case "n":
                    return "n.";
                default:
                    return "Genusfehler";
            }
        }

        /// <summary>
        /// Writes a note (comment, conjugation, ...) as a single cell table.
        ///
        /// The table is created with AddTable and then handed over to
        /// InsertTable. Only that combination adds the "Table Grid" style to
        /// styles.xml, and only that style carries the border of the box.
        /// </summary>
        private static void InsertNoteTable(DocX doc, string text, Formatting formatting)
        {
            var table = doc.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs.First().InsertText(text, false, formatting);
            table.AutoFit = AutoFit.Window;
            doc.InsertTable(table);
        }

        /// <summary>Writes the phrases that are not assigned to any entry.</summary>
        private static void WriteUnassignedPhrases(string file, List<Entry> phrases)
        {
            var doc = DocX.Create(file);
            var body = doc.Xml;

            body.Add(CreateParagraph(CreateRun(Environment.NewLine, PhraseFontSize, false, false)));
            body.Add(CreateParagraph(CreateRun("Unzugeordnete Phrasen:", EntryFontSize, false, false)));
            body.Add(CreateParagraph(CreateRun(Environment.NewLine, PhraseFontSize, false, false)));

            foreach (var phrase in phrases)
            {
                var paragraph = CreateParagraph();
                paragraph.Add(CreateRun(phrase.Ostfriesisch, PhraseFontSize, false, false));
                paragraph.Add(CreateRun(" ", PhraseFontSize, false, false));
                paragraph.Add(CreateRun(phrase.Deutsch, PhraseFontSize, false, true));
                body.Add(paragraph);
            }

            ApplyLayout(doc);
            doc.Save();
            doc.Dispose();
        }

        /// <summary>
        /// Applies the layout: two columns per page and the page number in the
        /// bottom right corner. Runs for every part file and for the full document.
        /// </summary>
        private static void ApplyLayout(DocX doc)
        {
            var body = doc.Xml;
            if (body == null)
            {
                return;
            }

            // Page number in the bottom right corner of the footer.
            doc.AddFooters();
            var footer = doc.Footers.odd;
            if (footer != null)
            {
                var footerParagraph = footer.Paragraphs.Count > 0
                    ? footer.Paragraphs[0]
                    : footer.InsertParagraph();
                footerParagraph.Alignment = Alignment.right;
                if (!footerParagraph.Xml.Descendants(W + "fldSimple").Any())
                {
                    footerParagraph.AppendPageNumber(PageNumberFormat.normal);
                }
            }

            var defaultFooterReference = SectionProperties(body)
                .SelectMany(sectionProperties => sectionProperties.Elements(W + "footerReference"))
                .FirstOrDefault(reference => (string)reference.Attribute(W + "type") == "default");

            foreach (var sectionProperties in SectionProperties(body))
            {
                // Page margins: 2 cm top and bottom, 2.54 cm left and right.
                var pageMargins = sectionProperties.Element(W + "pgMar");
                if (pageMargins == null)
                {
                    pageMargins = new XElement(W + "pgMar");
                    var pageSize = sectionProperties.Element(W + "pgSz");
                    if (pageSize != null)
                    {
                        pageSize.AddAfterSelf(pageMargins);
                    }
                    else
                    {
                        sectionProperties.AddFirst(pageMargins);
                    }
                }
                pageMargins.SetAttributeValue(W + "top", MarginTopBottom);
                pageMargins.SetAttributeValue(W + "bottom", MarginTopBottom);
                pageMargins.SetAttributeValue(W + "left", MarginLeftRight);
                pageMargins.SetAttributeValue(W + "right", MarginLeftRight);
                if (pageMargins.Attribute(W + "header") == null)
                {
                    pageMargins.SetAttributeValue(W + "header", 708);
                }
                if (pageMargins.Attribute(W + "footer") == null)
                {
                    pageMargins.SetAttributeValue(W + "footer", 708);
                }
                if (pageMargins.Attribute(W + "gutter") == null)
                {
                    pageMargins.SetAttributeValue(W + "gutter", 0);
                }

                // Two columns. The column spacing of the template is kept.
                var columns = sectionProperties.Element(W + "cols");
                if (columns == null)
                {
                    columns = new XElement(W + "cols");
                    sectionProperties.Add(columns);
                }
                columns.SetAttributeValue(W + "num", "2");
                if (columns.Attribute(W + "space") == null)
                {
                    columns.SetAttributeValue(W + "space", "708");
                }

                // Word only inherits a footer from the previous section. Every
                // section therefore points to the same footer explicitly so the
                // page number appears on every page.
                if (defaultFooterReference != null && !sectionProperties.Elements(W + "footerReference").Any())
                {
                    sectionProperties.AddFirst(new XElement(defaultFooterReference));
                }
            }

            // The format expects header and footer references at the beginning of
            // the section properties. The library appends them at the end, so
            // they are moved to the front here.
            foreach (var sectionProperties in SectionProperties(body))
            {
                var references = sectionProperties.Elements()
                    .Where(element => element.Name == W + "headerReference" || element.Name == W + "footerReference")
                    .ToList();
                if (references.Count == 0)
                {
                    continue;
                }

                var first = sectionProperties.Elements().First();
                if (references.Contains(first))
                {
                    continue;
                }

                foreach (var reference in references)
                {
                    reference.Remove();
                }
                sectionProperties.AddFirst(references);
            }

            // The section properties of the closing section belong to the end of
            // the body. The template puts them first, which is not what the
            // format expects, so they are moved to their proper place here.
            var bodySectionProperties = body.Elements(W + "sectPr").LastOrDefault();
            if (bodySectionProperties != null && body.Elements().Last() != bodySectionProperties)
            {
                bodySectionProperties.Remove();
                body.Add(bodySectionProperties);
            }
        }

        /// <summary>
        /// Returns the section properties of a document. They are either a direct
        /// child of the body (the closing section) or sit in the paragraph that
        /// ends a section. Walking only those two places is much faster than
        /// searching the whole document with Descendants.
        /// </summary>
        private static IEnumerable<XElement> SectionProperties(XElement body)
        {
            foreach (var child in body.Elements())
            {
                if (child.Name == W + "sectPr")
                {
                    yield return child;
                    continue;
                }

                if (child.Name != W + "p")
                {
                    continue;
                }

                var sectionProperties = child.Element(W + "pPr")?.Element(W + "sectPr");
                if (sectionProperties != null)
                {
                    yield return sectionProperties;
                }
            }
        }

        /// <summary>
        /// Merges the part files into one document. The content is taken on XML
        /// level, so the part files are parsed once instead of being copied as
        /// complete documents.
        ///
        /// The first part file is used as the base of the new document. That way
        /// the merged file carries the same styles as the part files, above all
        /// the "Table Grid" style that draws the border of the note boxes.
        /// </summary>
        private static void MergeFiles(string targetFile, IEnumerable<string> files)
        {
            var fileList = files.ToList();
            if (fileList.Count == 0)
            {
                return;
            }

            var doc = DocX.Load(fileList[0]);
            var body = doc.Xml;

            // The content of the first file is added again below.
            foreach (var node in body.Elements().Where(element => element.Name != W + "sectPr").ToList())
            {
                node.Remove();
            }

            var batchSize = Math.Max(1, Environment.ProcessorCount);

            // Read ahead in small batches: parsing is the expensive part and can
            // run in parallel, while assembling the document order has to be
            // sequential.
            for (var start = 0; start < fileList.Count; start += batchSize)
            {
                var end = Math.Min(start + batchSize, fileList.Count);
                var batch = new XElement[end - start];
                Parallel.For(start, end, index => batch[index - start] = ReadDocumentBody(fileList[index]));

                for (var index = start; index < end; index++)
                {
                    var sourceBody = batch[index - start];
                    if (sourceBody == null)
                    {
                        continue;
                    }

                    var sourceSectionProperties = sourceBody.Elements(W + "sectPr").LastOrDefault();

                    foreach (var node in sourceBody.Elements().Where(element => element != sourceSectionProperties).ToList())
                    {
                        body.Add(node);
                    }

                    // Section break with the page settings of the part file. The
                    // footer references are only valid inside the part file, so they
                    // are dropped here and set again by ApplyLayout. The last part
                    // file ends in the closing section, so it needs no break.
                    if (sourceSectionProperties != null && index < fileList.Count - 1)
                    {
                        var sectionBreak = new XElement(sourceSectionProperties);
                        sectionBreak.Elements(W + "headerReference").Remove();
                        sectionBreak.Elements(W + "footerReference").Remove();
                        body.Add(new XElement(W + "p", new XElement(W + "pPr", sectionBreak)));
                    }
                }
            }

            ApplyLayout(doc);
            doc.SaveAs(targetFile);
            doc.Dispose();
        }

        /// <summary>Reads the content of word/document.xml of a docx file.</summary>
        private static XElement ReadDocumentBody(string file)
        {
            using (var archive = ZipFile.OpenRead(file))
            {
                var entry = archive.GetEntry("word/document.xml");
                if (entry == null)
                {
                    return null;
                }
                using (var stream = entry.Open())
                {
                    var document = XElement.Load(stream);
                    return document.Element(W + "body");
                }
            }
        }

        /// <summary>Creates an empty paragraph element.</summary>
        private static XElement CreateParagraph()
        {
            return new XElement(W + "p", new XElement(W + "pPr"));
        }

        private static XElement CreateParagraph(XElement run)
        {
            var paragraph = CreateParagraph();
            paragraph.Add(run);
            return paragraph;
        }

        /// <summary>Creates a text run with the given font size and style.</summary>
        private static XElement CreateRun(string text, double size, bool bold, bool italic)
        {
            var runProperties = new XElement(W + "rPr",
                new XElement(W + "rFonts",
                    new XAttribute(W + "ascii", FontFamily),
                    new XAttribute(W + "hAnsi", FontFamily),
                    new XAttribute(W + "cs", FontFamily)));

            if (bold)
            {
                runProperties.Add(new XElement(W + "b"));
            }
            if (italic)
            {
                runProperties.Add(new XElement(W + "i"));
            }

            var halfPoints = ((int)Math.Round(size * 2)).ToString(CultureInfo.InvariantCulture);
            runProperties.Add(new XElement(W + "sz", new XAttribute(W + "val", halfPoints)));
            runProperties.Add(new XElement(W + "szCs", new XAttribute(W + "val", halfPoints)));

            var run = new XElement(W + "r", runProperties);

            // A line break inside the text becomes a real line break in Word.
            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    run.Add(new XElement(W + "br"));
                }
                if (lines[i].Length == 0)
                {
                    continue;
                }

                var textNode = new XElement(W + "t", lines[i]);
                if (lines[i].Length != lines[i].Trim().Length)
                {
                    textNode.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                }
                run.Add(textNode);
            }

            return run;
        }

        private static void Log(string message)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }
    }
}
