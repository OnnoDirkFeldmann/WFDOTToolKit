using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text.RegularExpressions;

namespace WFDOTToolKit
{
    public class WFDOTToDic
    {
        public static string Todic()
        {
            string dictstring;
            dictstring = "";
            SQLiteConnection connection = new SQLiteConnection();
            connection.ConnectionString = @"Data Source=D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";
            connection.Open();
            string sqlCom = "SELECT * FROM WB";
            SQLiteCommand scdCommand = new SQLiteCommand(sqlCom, connection);
            SQLiteDataReader reader = scdCommand.ExecuteReader();
            List<string> dicw = new List<string>();
            while (reader.Read())
            {
                string ofrs;
                ofrs = (string)reader["Ostfriesisch"];
                ofrs = ofrs.Replace("!", " ");
                ofrs = ofrs.Replace("?", " ");
                ofrs = ofrs.Replace("\"", " ");
                ofrs = ofrs.Replace(" - ", " ");
                ofrs = ofrs.Replace("\n", " ");
                ofrs = ofrs.Replace("\r", " ");
                dicw.AddRange(ofrs.Split(" "));
                string plural;
                plural = (string)reader["Plural"];
                if (!plural.Equals("-"))
                {
                    plural = Regex.Replace(plural, "\\((.+?)\\)", "");
                    plural = plural.Replace(",", " ");
                    plural = plural.Replace("\n", " ");
                    plural = plural.Replace("\r", " ");
                    dicw.AddRange(plural.Split(" "));
                }
                string komparation;
                komparation = (string)reader["Komparation"];
                if (!komparation.Equals("-"))
                {
                    komparation = Regex.Replace(komparation, "\\((.+?)\\)", "");
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
                   
                    dicw.AddRange(komparation.Split(" "));
                }
                string konjugation;
                konjugation = (string)reader["Konjugation"];
                if (!konjugation.Equals("-"))
                {
                    konjugation = Regex.Replace(konjugation, "\\((.+?)\\)", "");
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
                    dicw.AddRange(konjugation.Split(" "));
                }
            }
            dicw = dicw.Distinct().ToList();
            dicw.Sort();
            for (int i = 0; i < dicw.Count; i++)
            {
                dictstring = dictstring + dicw[i] + "\r\n";
            }

            //alles Trimmen
            return dictstring;
        }
    }
}
