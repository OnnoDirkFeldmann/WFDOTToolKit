using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WFDOTToolKit
{
    /// <summary>
    /// Builds the word list for the FRS spell checker (frs_FRS.txt).
    ///
    /// Format of the generated file (same as the reference file):
    ///   * one entry per line: "word;1"
    ///   * UTF-8 without byte order mark
    ///   * CRLF as line separator, the last line has no line break
    ///   * sorted by the complete line, so "word-;1" comes before "word;1"
    ///
    /// The words are the same as for the dic file: headword, plural, comparison
    /// and conjugation of every dictionary entry.
    /// </summary>
    public class WFDOTToFrs
    {
        /// <summary>Path of the dictionary database.</summary>
        private const string DatabasePath = @"D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";

        /// <summary>Default target of the generated file.</summary>
        public const string DefaultTarget = @"C:\Users\Neronno\Desktop\frs_FRS.txt";

        /// <summary>Flag that the FRS spell checker uses for a normal word.</summary>
        private const string WordFlag = ";1";

        /// <summary>Generates the FRS word list and writes it to <paramref name="target"/>.</summary>
        public static void WriteFrs(string target)
        {
            var words = CollectWords();
            var lines = new List<string>(words.Count);
            foreach (var word in words)
            {
                lines.Add(word + WordFlag);
            }

            // Sorted by the whole line. That is what the reference file does:
            // "Auerk-;1" comes before "Auerk:;1" and "Auerk;1" comes last.
            lines.Sort(StringComparer.Ordinal);

            var content = string.Join("\r\n", lines);
            var directory = Path.GetDirectoryName(target);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // UTF-8 without byte order mark.
            File.WriteAllText(target, content, new UTF8Encoding(false));

            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {lines.Count} Wörter nach {target} geschrieben.");
        }

        /// <summary>Collects all usable words of the dictionary.</summary>
        private static List<string> CollectWords()
        {
            var words = new HashSet<string>(StringComparer.Ordinal);

            using (var connection = new SQLiteConnection("Data Source=" + DatabasePath))
            {
                connection.Open();
                using (var command = new SQLiteCommand("SELECT * FROM WB", connection))
                using (var reader = command.ExecuteReader())
                {
                    var ordinalOstfriesisch = reader.GetOrdinal("Ostfriesisch");
                    var ordinalPlural = reader.GetOrdinal("Plural");
                    var ordinalKomparation = reader.GetOrdinal("Komparation");
                    var ordinalKonjugation = reader.GetOrdinal("Konjugation");

                    while (reader.Read())
                    {
                        AddWords(words, CleanHeadword(reader.GetString(ordinalOstfriesisch)));

                        var plural = reader.GetString(ordinalPlural);
                        if (plural != "-")
                        {
                            AddWords(words, RemoveBrackets(plural).Replace(",", " "));
                        }

                        var komparation = reader.GetString(ordinalKomparation);
                        if (komparation != "-")
                        {
                            AddWords(words, CleanKomparation(komparation));
                        }

                        var konjugation = reader.GetString(ordinalKonjugation);
                        if (konjugation != "-")
                        {
                            AddWords(words, CleanKonjugation(konjugation));
                        }
                    }
                }
            }

            var result = words.ToList();
            result.Sort(StringComparer.Ordinal);
            return result;
        }

        /// <summary>Splits a text into words and adds the usable ones.</summary>
        private static void AddWords(HashSet<string> words, string text)
        {
            var parts = text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (IsUsableWord(part))
                {
                    words.Add(part);
                }
            }
        }

        /// <summary>
        /// Cleans the headword: quotation marks, question marks and exclamation
        /// marks separate the words instead of being part of them.
        /// </summary>
        private static string CleanHeadword(string text)
        {
            text = text.Replace("!", " ");
            text = text.Replace("?", " ");
            text = text.Replace("\"", " ");
            text = text.Replace(" - ", " ");
            text = text.Replace("\n", " ");
            text = text.Replace("\r", " ");
            return text;
        }

        /// <summary>Removes the bracketed explanations, e.g. "(prädikativ)".</summary>
        private static string RemoveBrackets(string text)
        {
            return Regex.Replace(text, "\\((.+?)\\)", "");
        }

        /// <summary>Removes the labels of the comparison forms, keeps the forms.</summary>
        private static string CleanKomparation(string text)
        {
            text = RemoveBrackets(text);
            text = text.Replace("stark<br/>Positiv: ", " ");
            text = text.Replace("schwach<br/>Positiv: ", " ");
            text = text.Replace("<br/>Komparativ (prädikativ): ", " ");
            text = text.Replace("<br/>Komparativ (attributiv): ", " ");
            text = text.Replace("<br/>Superlativ : ", " ");
            text = text.Replace("<br/>Attributiv (bestimmter Artikel/unbestimmter Artikel m./f.): ", " ");
            text = text.Replace("<br/>Attributiv (unbestimmter Artikel n.): ", " ");
            text = text.Replace("<br/>Elativ: ", " ");
            text = text.Replace("<br/>", " ");
            text = text.Replace("/", " ");
            return text;
        }

        /// <summary>Removes the labels of the conjugated forms, keeps the forms.</summary>
        private static string CleanKonjugation(string text)
        {
            text = RemoveBrackets(text);
            text = text.Replace("stark<br/>Inf.: ", " ");
            text = text.Replace("schwach<br/>Inf.: ", " ");
            text = text.Replace("stark (Praet.-Praes.)<br/>Inf.: ", " ");
            text = text.Replace("<br/>Imp: ", " ");
            text = text.Replace("<br/>Präs.: ", " ");
            text = text.Replace("<br/>Prät.: ", " ");
            text = text.Replace("<br/>Part.I: ", " ");
            text = text.Replace("<br/>Part.II: ", " ");
            text = text.Replace("<br/>", " ");
            text = text.Replace(",", " ");
            text = text.Replace("/", " ");
            return text;
        }

        /// <summary>
        /// Decides whether a token becomes an entry of the word list. Pure numbers
        /// and plain punctuation are dropped, everything that carries a letter is
        /// kept. Typographic characters such as the ellipsis are kept as well.
        /// </summary>
        private static bool IsUsableWord(string word)
        {
            if (word.Length == 0)
            {
                return false;
            }

            if (word.Any(char.IsLetter))
            {
                return true;
            }

            return word.Any(character => character > 127) && !word.Any(char.IsDigit);
        }
    }
}
