using Novacode;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;

namespace WFDOTToolKit
{
    public class WFDOTToDocx
    {
        public static void ToDocx(bool fullversion, string path, bool testMode)
        {
            Directory.Delete(path, true);
            Directory.CreateDirectory(path);
            SQLiteConnection connection = new SQLiteConnection();
            connection.ConnectionString = @"Data Source=D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";
            connection.Open();
            var sql1 = testMode ? "SELECT * FROM WB WHERE Ostfriesisch Like 'a%'" : "SELECT * FROM WB";
            string sqlCom = sql1;
            SQLiteCommand scdCommand = new SQLiteCommand(sqlCom, connection);
            SQLiteDataReader reader = scdCommand.ExecuteReader();
            WFDOT wfdot = new WFDOT();
            wfdot._WFDOT.Load(reader);
            reader.Close();
            connection.Close();

            var wordFormat = new Formatting();
            wordFormat.FontFamily = new Font("Segoe UI");
            wordFormat.Size = 9D;

            var wordFormatBold = new Formatting();
            wordFormatBold.FontFamily = new Font("Segoe UI");
            wordFormatBold.Size = 9D;
            wordFormatBold.Bold = true;

            var textFormatBold = new Formatting();
            textFormatBold.FontFamily = new Font("Segoe UI");
            textFormatBold.Size = 8D;
            textFormatBold.Bold = true;

            var textFormat = new Formatting();
            textFormat.FontFamily = new Font("Segoe UI");
            textFormat.Size = 8D;

            var textFormatItalic = new Formatting();
            textFormatItalic.FontFamily = new Font("Segoe UI");
            textFormatItalic.Size = 8D;
            textFormatItalic.Italic = true;

            //Oostfräisk woorden hóólen
            var files = new List<string>();
            var i = 0;
            var docIndex = 0;
            var entries = wfdot._WFDOT.Where(x => x.Wortart != "Phrase").ToList();
            var fileOne = $@"{path}\out{docIndex}.docx";
            files.Add(fileOne);
            var doc = DocX.Create(fileOne);
            foreach (var row in entries)
            {
                if (i % 1000 == 0 && i != 0)
                {
                    docIndex++;
                    doc.Save();
                    var fileNext = $@"{path}\out{docIndex}.docx";
                    files.Add(fileNext);
                    doc = DocX.Create(fileNext);
                    Console.WriteLine(row.Ostfriesisch);
                }
                i++;

                var paragraph = doc.InsertParagraph("", false, wordFormat);

                paragraph.InsertText(row.Ostfriesisch, false, wordFormatBold);

                string wortartstring = "Wortartfehler";

                if (row.Wortart != "-")
                {
                    switch (row.Wortart)
                    {
                        case "Abkürzung":
                            wortartstring = "-";
                            break;
                        case "Adjektiv":
                            wortartstring = "-";
                            break;
                        case "Adverb":
                            wortartstring = "-";
                            break;
                        case "Artikel":
                            wortartstring = "-";
                            break;
                        case "Ausruf":
                            wortartstring = "-";
                            break;
                        case "Flexionsform":
                            wortartstring = "-";
                            break;
                        case "Interrogativpronomen":
                            wortartstring = "-";
                            break;
                        case "Konjunktion":
                            wortartstring = "-";
                            break;
                        case "Nachsilbe":
                            wortartstring = "-";
                            break;
                        case "Name":
                            wortartstring = "-";
                            break;
                        case "Numeral":
                            wortartstring = "-";
                            break;
                        case "Ortsname":
                            wortartstring = "-";
                            break;
                        case "Partikel":
                            wortartstring = "-";
                            break;
                        case "Partizip":
                            wortartstring = "-";
                            break;
                        case "Phrase":
                            wortartstring = "-";
                            break;
                        case "Pronomen":
                            wortartstring = "-";
                            break;
                        case "Pronominaladverb":
                            wortartstring = "-";
                            break;
                        case "Substantiv":
                            wortartstring = "-";
                            break;
                        case "Verb":
                            wortartstring = "-";
                            break;
                        case "Vorsilbe":
                            wortartstring = "-";
                            break;
                        case "Zwischensilbe":
                            wortartstring = "-";
                            break;
                    }
                }

                if (row.Artikel != "-" || row.Genus != "-" || row.Plural != "-" || (row.Wortart != "-" && wortartstring != "-"))
                {
                    paragraph.InsertText(" [", false, wordFormat);
                }

                if (row.Artikel != "-")
                {
                    paragraph.InsertText(row.Artikel, false, wordFormat);
                }

                if (wortartstring != "-")
                {
                    if (row.Artikel != "-") paragraph.InsertText(", ", false, wordFormat);
                    paragraph.InsertText(wortartstring, false, wordFormat);
                }

                if (row.Genus != "-")
                {
                    string genusstring = "Genusfehler";
                    switch (row.Genus)
                    {
                        case "m":
                            genusstring = "m.";
                            break;
                        case "f":
                            genusstring = "f.";
                            break;
                        case "n":
                            genusstring = "n.";
                            break;
                    }
                    if (row.Artikel != "-" || wortartstring != "-") paragraph.InsertText(", ", false, wordFormat);
                    paragraph.InsertText(genusstring, false, wordFormat);
                }

                if (row.Plural != "-")
                {
                    if (row.Artikel != "-" || wortartstring != "-" || row.Plural != "-") paragraph.InsertText(", ", false, wordFormat);
                    paragraph.InsertText(row.Plural, false, wordFormat);
                }

                if (row.Artikel != "-" || row.Genus != "-" || row.Plural != "-" || (row.Wortart != "-" && wortartstring != "-"))
                {
                    paragraph.InsertText("]", false, wordFormat);
                }

                paragraph.InsertText(" ", false, wordFormat);
                paragraph.InsertText(row.Deutsch, false, wordFormat);

                if (row.Nebenformen != "-")
                {
                    paragraph.InsertText(Environment.NewLine, false, wordFormat);
                    paragraph.InsertText($"[NF: {row.Nebenformen}]", false, wordFormat);
                }
                if (row.Standardform != "-")
                {
                    paragraph.InsertText(Environment.NewLine, false, wordFormat);
                    paragraph.InsertText($"[{row.Standardform}]", false, wordFormat);
                }

                if (fullversion)
                {
                    if (row.Komparation != "-")
                    {
                        var komparation = row.Komparation;
                        komparation = komparation.Substring(0, komparation.Length - 5);
                        komparation = komparation.Replace("<br/>", "; ");
                        var table = doc.AddTable(1, 1);
                        table.Rows[0].Cells[0].Paragraphs.First().InsertText(komparation, false, textFormat);
                        table.AutoFit = AutoFit.Window;
                        doc.InsertTable(table);
                    }
                    if (row.Konjugation != "-")
                    {
                        var konjugation = row.Konjugation;
                        konjugation = konjugation.Substring(0, konjugation.Length - 5);
                        konjugation = konjugation.Replace("<br/>", "; ");
                        var table = doc.AddTable(1, 1);
                        table.Rows[0].Cells[0].Paragraphs.First().InsertText(konjugation, false, textFormat);
                        table.AutoFit = AutoFit.Window;
                        doc.InsertTable(table);
                    }
                    if (row.Kommentar != "-")
                    {
                        var table = doc.AddTable(1, 1);
                        table.Rows[0].Cells[0].Paragraphs.First().InsertText(row.Kommentar, false, textFormat);
                        table.AutoFit = AutoFit.Window;
                        doc.InsertTable(table);
                    }
                    if (row.Rezept != "-")
                    {
                        var rezept = row.Rezept;
                        rezept = rezept.Replace("<br/>", " ");
                        var table = doc.AddTable(1, 1);
                        table.Rows[0].Cells[0].Paragraphs.First().InsertText(rezept, false, textFormat);
                        table.AutoFit = AutoFit.Window;
                        doc.InsertTable(table);
                    }
                }
                // Phrasen anreichern
                string zuordnung = row.Nummer == "-" ? row.Ostfriesisch : row.Ostfriesisch + "=" + row.Nummer;
                var phrases = wfdot._WFDOT.Where(x => x.Wortart == "Phrase" && x.Zuordnung == zuordnung).ToList();
                if (phrases.Any())
                {
                    var phraseParagraph = doc.InsertParagraph("", false, textFormat);
                    for (int j = 0; j < phrases.Count; j++)
                    {
                        WFDOT.WFDOTRow row2 = phrases[j];
                        phraseParagraph.InsertText(row2.Ostfriesisch, false, textFormatBold);
                        phraseParagraph.InsertText(" ", false, textFormat);
                        phraseParagraph.InsertText(row2.Deutsch, false, textFormatItalic);
                        if (j != phrases.Count - 1) phraseParagraph.InsertText(" ", false, textFormat);
                    }
                }
            }

            doc.Save();

            var ohneZuordnung = wfdot._WFDOT.Where(x => x.Wortart == "Phrase" && x.Zuordnung == "-").ToList();
            if (ohneZuordnung.Any())
            {
                //fróósen sünner tauörnen
                var fileOhneZuordnung = $@"{path}\outOhneZu.docx";
                files.Add(fileOhneZuordnung);
                doc = DocX.Create(fileOhneZuordnung);
                doc.InsertParagraph(Environment.NewLine, false, textFormat);
                doc.InsertParagraph("Unzugeordnete Phrasen:", false, wordFormat);
                doc.InsertParagraph(Environment.NewLine, false, textFormat);
                foreach (var row3 in ohneZuordnung)
                {
                    var phraseParagraph = doc.InsertParagraph("", false, textFormat);
                    phraseParagraph.InsertText(row3.Ostfriesisch, false, textFormat);
                    phraseParagraph.InsertText(" ", false, textFormat);
                    phraseParagraph.InsertText(row3.Deutsch, false, textFormatItalic);
                }
                doc.Save();
            }

            var fullDoc = DocX.Create($@"{path}\full.docx");
            foreach (var document in files)
            {
                var partDoc = DocX.Load(document);
                fullDoc.InsertDocument(partDoc, true);
            }
            fullDoc.Save();
        }
    }
}
