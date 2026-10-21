using System;
using System.IO;
using System.Text;

namespace WFDOTToolKit
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            // Non-interactive import: WFDOTToolKit.exe <excel> <db>
            if (args.Length == 2 && File.Exists(args[0]))
            {
                WFDOTImport.Import(args[0], args[1]);
                return;
            }

            // Non-interactive cleanup: WFDOTToolKit.exe cleanup <excel>
            if (args.Length == 2 && args[0].Equals("cleanup", StringComparison.OrdinalIgnoreCase) && File.Exists(args[1]))
            {
                WFDOTImport.Cleanup(args[1]);
                return;
            }

            string line;
            Console.WriteLine("=== Ausgabe ===");
            Console.WriteLine(" 1=html                  2=html (full)          3=docx");
            Console.WriteLine(" 4=docx-full             5=docx-full and pdf    6=docx-full testpage");
            Console.WriteLine("=== Wörterbuch-Datei ===");
            Console.WriteLine(" 7=dic                  8=frs_FRS");
            Console.WriteLine("=== Prüfen ===");
            Console.WriteLine(" 9=counterpart without index   10=wrong assignments");
            Console.WriteLine("=== Import ===");
            Console.WriteLine("11=import excel         12=excel cleanup");
            line = Console.ReadLine();
            switch (line)
            {
                case "1":
                    string html;
                    html = WFDOTToHtml.Tohtml(false);
                    File.WriteAllText(@"C:\Users\Neronno\Desktop\out.html", html);
                    break;
                case "2":
                    string htmlFull;
                    htmlFull = WFDOTToHtml.Tohtml(true);
                    File.WriteAllText(@"C:\Users\Neronno\Desktop\outfull.html", htmlFull);
                    break;
                case "3":
                    WFDOTToDocx.ToDocx(false, @"C:\Users\Neronno\Desktop\gen", false);
                    break;
                case "4":
                    WFDOTToDocx.ToDocx(true, @"C:\Users\Neronno\Desktop\gen", false);
                    break;
                case "5":
                    string genPath = @"C:\Users\Neronno\Desktop\gen";
                    WFDOTToDocx.ToDocx(true, genPath, false);
                    try
                    {
                        WFDOTToPdf.ExportToPdf(
                            WFDOTToDocx.GetMainFile(genPath),
                            Path.Combine(genPath, "WFDOT.pdf"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"FEHLER beim PDF-Export: {ex.Message}");
                    }
                    break;
                case "6":
                    WFDOTToDocx.ToDocx(true, @"C:\Users\Neronno\Desktop\gen", true);
                    break;
                case "7":
                    WFDOTToDic.WriteDic(@"C:\Users\Neronno\Desktop\frs.dic");
                    break;
                case "8":
                    WFDOTToFrs.WriteFrs(WFDOTToFrs.DefaultTarget);
                    break;
                case "9":
                    WFDOTChecker.DublicatesWithoutIndex();
                    break;
                case "10":
                    WFDOTChecker.ShowWrongReference();
                    break;
                case "11":
                    WFDOTImport.Import();
                    break;
                case "12":
                    WFDOTImport.Cleanup();
                    break;
            }
        }
    }
}

