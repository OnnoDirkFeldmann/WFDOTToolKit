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
    /// Builds the word list of the spell checker (frs.dic).
    ///
    /// The list holds the headword, the plural, the comparison forms and the
    /// conjugated forms of every dictionary entry, one word per line.
    ///
    /// The labels of the forms are removed, so only the forms themselves end up
    /// in the list. The source file of this generator was written in CP1252 and
    /// was converted to UTF-8 in the meantime, the generated file stays the
    /// same.
    /// </summary>
    public class WFDOTToDic
    {
        /// <summary>Path of the dictionary database.</summary>
        private const string DatabasePath = @"D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";

        /// <summary>Removes the explanations in brackets, e.g. "(prädikativ)".</summary>
        private static readonly Regex Brackets = new Regex("\\((.+?)\\)", RegexOptions.Compiled);

        /// <summary>Returns the whole word list as one string.</summary>
        public static string Todic()
        {
            var words = CollectWords();

            var builder = new StringBuilder(words.Count * 12);
            foreach (var word in words)
            {
                builder.Append(word);
                builder.Append("\r\n");
            }
            return builder.ToString();
        }

        /// <summary>
        /// Writes the word list into the target file as UTF-16 with byte order
        /// mark. The words are streamed, so the big text is never held in memory.
        /// </summary>
        public static void WriteDic(string path)
        {
            var words = CollectWords();

            using (var writer = new StreamWriter(path, false, Encoding.Unicode))
            {
                writer.NewLine = "\r\n";
                foreach (var word in words)
                {
                    writer.Write(word);
                    writer.Write("\r\n");
                }
            }
        }

        /// <summary>
        /// Reads the dictionary and returns all words, sorted and without
        /// duplicates.
        /// </summary>
        private static List<string> CollectWords()
        {
            var words = new List<string>(400000);

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
                        var headword = reader.GetString(ordinalOstfriesisch);
                        headword = headword.Replace("!", " ");
                        headword = headword.Replace("?", " ");
                        headword = headword.Replace("\"", " ");
                        headword = headword.Replace(" - ", " ");
                        headword = headword.Replace("\n", " ");
                        headword = headword.Replace("\r", " ");
                        words.AddRange(headword.Split(' '));

                        var plural = reader.GetString(ordinalPlural);
                        if (!plural.Equals("-"))
                        {
                            plural = Brackets.Replace(plural, "");
                            plural = plural.Replace(",", " ");
                            plural = plural.Replace("\n", " ");
                            plural = plural.Replace("\r", " ");
                            words.AddRange(plural.Split(' '));
                        }

                        var komparation = reader.GetString(ordinalKomparation);
                        if (!komparation.Equals("-"))
                        {
                            komparation = Brackets.Replace(komparation, "");
                            komparation = komparation.Replace("stark<br/>Positiv: ", " ");
                            komparation = komparation.Replace("schwach<br/>Positiv: ", " ");
                            komparation = komparation.Replace("<br/>Komparativ (prädikativ): ", " ");
                            komparation = komparation.Replace("<br/>Komparativ (attributiv): ", " ");
                            komparation = komparation.Replace("<br/>Superlativ : ", " ");
                            komparation = komparation.Replace("<br/>Attributiv (bestimmter Artikel/unbestimmter Artikel m./f.): ", " ");
                            komparation = komparation.Replace("<br/>Attributiv (unbestimmter Artikel n.): ", " ");
                            komparation = komparation.Replace("<br/>Elativ: ", " ");
                            komparation = komparation.Replace("<br/>", " ");
                            komparation = komparation.Replace("/", " ");
                            komparation = komparation.Replace("\n", " ");
                            komparation = komparation.Replace("\r", " ");
                            words.AddRange(komparation.Split(' '));
                        }

                        var konjugation = reader.GetString(ordinalKonjugation);
                        if (!konjugation.Equals("-"))
                        {
                            konjugation = Brackets.Replace(konjugation, "");
                            konjugation = konjugation.Replace("stark<br/>Inf.: ", " ");
                            konjugation = konjugation.Replace("schwach<br/>Inf.: ", " ");
                            konjugation = konjugation.Replace("stark (Praet.-Praes.)<br/>Inf.: ", " ");
                            konjugation = konjugation.Replace("<br/>Imp: ", " ");
                            konjugation = konjugation.Replace("<br/>Präs.: ", " ");
                            konjugation = konjugation.Replace("<br/>Prät.: ", " ");
                            konjugation = konjugation.Replace("<br/>Part.I: ", " ");
                            konjugation = konjugation.Replace("<br/>Part.II: ", " ");
                            konjugation = konjugation.Replace("<br/>", " ");
                            konjugation = konjugation.Replace(",", " ");
                            konjugation = konjugation.Replace("/", " ");
                            konjugation = konjugation.Replace("\n", " ");
                            konjugation = konjugation.Replace("\r", " ");
                            words.AddRange(konjugation.Split(' '));
                        }
                    }
                }
            }

            var result = words.Distinct().ToList();
            result.Sort();
            return result;
        }
    }
}
