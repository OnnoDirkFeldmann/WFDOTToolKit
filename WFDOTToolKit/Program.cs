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
            Console.WriteLine("1=html 2=html (full) 3=docx 4=docx-full testpage 5=docx-full 6=dic 7=counterpart without index 8=wrong assignments 9=import excel 10=excel cleanup");
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
                    WFDOTToDocx.ToDocx(true, @"C:\Users\Neronno\Desktop\gen", true);
                    break;
                case "5":
                    WFDOTToDocx.ToDocx(true, @"C:\Users\Neronno\Desktop\gen", false);
                    break;
                case "6":
                    string dic;
                    dic = WFDOTToDic.Todic();
                    File.WriteAllText(@"C:\Users\Neronno\Desktop\frs.dic", dic, Encoding.Unicode);
                    break;
                case "7":
                    WFDOTChecker.DublicatesWithoutIndex();
                    break;
                case "8":
                    WFDOTChecker.ShowWrongReference();
                    break;
                case "9":
                    WFDOTImport.Import();
                    break;
                case "10":
                    WFDOTImport.Cleanup();
                    break;
            }
        }
    }
}

