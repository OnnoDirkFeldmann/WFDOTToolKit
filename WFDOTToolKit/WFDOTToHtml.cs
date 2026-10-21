using System.Collections.Generic;
using System.Data.SQLite;

namespace WFDOTToolKit
{
    public class WFDOTToHtml
    {
        public static string Tohtml(bool fullversion)
        {
            SQLiteConnection connection = new SQLiteConnection();
            connection.ConnectionString = @"Data Source=D:\source\repos\WFDOTToolKit\WFDOTToolKit\WFDOT.db";
            connection.Open();

            string sqlCom = "SELECT * FROM WB WHERE Wortart != 'Phrase'";
            SQLiteCommand scdCommand = new SQLiteCommand(sqlCom, connection);
            SQLiteDataReader reader = scdCommand.ExecuteReader();

            //Ostfriesische Begriffe holen
            string ostfriesisch;
            string standardform;
            string konjugation;
            string nummer;
            string deutsch;
            string artikel;
            string plural;
            string temp;
            string wortart;
            string genus;
            string komparation;
            string nebenformen;
            string writestring;
            string kommentar;
            string autorKommentar;
            string rezept;
            string interferenz;
            writestring = "";
            List<string> list = new List<string>();
            List<string> ostfriesischewoerter = new List<string>();
            List<string> nummerwoerter = new List<string>();
            List<string> list2 = new List<string>();
            while (reader.Read())
            {
                ostfriesisch = (string)reader["Ostfriesisch"];
                ostfriesischewoerter.Add(ostfriesisch);
                ostfriesisch = "<span style=\"font-family:Verdana; font-size:12pt\"><b>" + ostfriesisch + "</b></span>";
                deutsch = (string)reader["Deutsch"];
                standardform = (string)reader["Standardform"];
                artikel = (string)reader["Artikel"];
                plural = (string)reader["Plural"];
                wortart = (string)reader["Wortart"];
                genus = (string)reader["Genus"];
                nebenformen = (string)reader["Nebenformen"];
                komparation = (string)reader["Komparation"];
                konjugation = (string)reader["Konjugation"];
                kommentar = (string)reader["Kommentar"];
                autorKommentar = (string)reader["Autorkommentar"];
                rezept = (string)reader["Rezept"];
                interferenz = (string)reader["Interferenz"];

                temp = ostfriesisch;

                string wortartstring = "Wortartfehler";
                if (wortart != "-")
                {
                    switch (wortart)
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



                if (artikel != "-" || genus != "-" || plural != "-" || (wortart != "-" && wortartstring != "-"))
                {
                    temp += "<span style=\"font-family:Verdana; font-size:12pt\"> [</span>";
                }

                if (wortartstring != "-")
                {
                    temp += "<span style=\"font-family:Verdana; font-size:12pt\">, " + wortartstring + "</span>";
                }


                if (artikel != "-")
                {
                    temp += "<span style=\"font-family:Verdana; font-size:12pt\">" + artikel + "</span>";
                }

                if (genus != "-")
                {
                    string genusstring = "Genusfehler";
                    switch (genus)
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
                    temp += "<span style=\"font-family:Verdana; font-size:12pt\">, " + genusstring + "</span>";
                }

                if (plural != "-")
                {
                    temp += "<span style=\"font-family:Verdana; font-size:12pt\">, " + plural + "</span>";
                }

                if (artikel != "-" || genus != "-" || plural != "-" || (wortart != "-" && wortartstring != "-"))
                {
                    temp += "<span style=\"font-family:Verdana; font-size:12pt\">]</span>";
                }

                temp += "<i><span style=\"font-family:Verdana; font-size:12pt\"> " + deutsch + "<br/></span></i>";
                if (nebenformen != "-")
                {
                    temp += "<span style=\"font-family:Verdana; font-size:11pt\">[NF: " + nebenformen + "]<br/></span>";
                }
                if (standardform != "-")
                {
                    temp += "<span style=\"font-family:Verdana; font-size:11pt\">[" + standardform + "]<br/></span>";
                }

                if (fullversion)
                {
                    if (komparation != "-")
                    {
                        komparation = komparation.Substring(0, komparation.Length - 5);
                        komparation = komparation.Replace("<br/>", ";");
                        temp += "<div style=\"border:1px solid black;\"><span style=\"font-family:Verdana; font-size:11pt\">" + komparation + "<br/></span></div>";
                    }
                    if (konjugation != "-")
                    {
                        konjugation = konjugation.Substring(0, konjugation.Length - 5);
                        konjugation = konjugation.Replace("<br/>", ";");
                        temp += "<div style=\"border:1px solid black;\"><span style=\"font-family:Verdana; font-size:11pt\">" + konjugation + "<br/></span></div>";
                    }
                    if (kommentar != "-")
                    {
                        temp += "<div style=\"border:1px solid black;\"><span style=\"font-family:Verdana; font-size:11pt\">" + kommentar + "<br/></span></div>";
                    }
                    if (autorKommentar != "-")
                    {
                        temp += "<div style=\"border:1px solid black;\"><span style=\"font-family:Verdana; font-size:11pt\">" + autorKommentar + "<br/></span></div>";
                    }
                    if (rezept != "-")
                    {
                        temp += "<div style=\"border:1px solid black;\"><span style=\"font-family:Verdana; font-size:11pt\">" + rezept + "<br/></span></div>";
                    }
                    if (interferenz != "-")
                    {
                        temp += "<div style=\"border:1px solid black;\"><span style=\"font-family:Verdana; font-size:11pt\">" + interferenz + "<br/></span></div>";
                    }
                }

                nummer = (string)reader["Index"];
                nummerwoerter.Add(nummer);
                list.Add(temp);
            }
            reader.Close();
            // Phrasen anreichern
            for (int i = 0; i < ostfriesischewoerter.Count; i++)
            {
                string final;
                final = list[i];
                SQLiteParameter ofrsprep = new SQLiteParameter("@ofrs");
                string sqlCom2 = "SELECT * FROM WB WHERE Wortart = 'Phrase' AND Zuordnung = @ofrs";
                SQLiteCommand scdCommand2 = new SQLiteCommand(sqlCom2, connection);
                if (nummerwoerter[i] == "-")
                {
                    ofrsprep.Value = ostfriesischewoerter[i];
                }
                else
                {
                    ofrsprep.Value = ostfriesischewoerter[i] + "=" + nummerwoerter[i];
                }
                scdCommand2.Parameters.Add(ofrsprep);
                scdCommand2.Prepare();
                SQLiteDataReader reader2 = scdCommand2.ExecuteReader();
                while (reader2.Read())
                {
                    ostfriesisch = (string)reader2["Ostfriesisch"];
                    deutsch = (string)reader2["Deutsch"];
                    deutsch = "<i>" + deutsch + "</i>";
                    temp = "<span style=\"font-family:Verdana; font-size:11pt\">" + ostfriesisch + " " + deutsch + "<br/></span>";
                    final = final + temp;
                }
                writestring = writestring + final;
                reader2.Close();
            }
            //Phrasen ohne Zuordnung
            string addition = "<span style=\"font-family:Verdana; font-size:11pt\"><br/><br/>Unzugeordnete Phrasen:<br/><br/><br/></span>";
            string sqlCom3 = "SELECT * FROM WB WHERE Wortart = 'Phrase' AND Zuordnung = '-'";
            SQLiteCommand scdCommand3 = new SQLiteCommand(sqlCom3, connection);
            scdCommand3.Prepare();
            SQLiteDataReader reader3 = scdCommand3.ExecuteReader();
            while (reader3.Read())
            {
                ostfriesisch = (string)reader3["Ostfriesisch"];
                deutsch = (string)reader3["Deutsch"];
                deutsch = "<i>" + deutsch + "</i>";
                temp = "<span style=\"font-family:Verdana; font-size:11pt\">" + ostfriesisch + " " + deutsch + "<br/></span>";
                addition = addition + temp;
            }
            reader3.Close();
            writestring = writestring + addition;
            writestring = "<!DOCTYPE html><html><body>" + writestring + "</body></html>";
            connection.Close();
            return writestring;
        }
    }
}
