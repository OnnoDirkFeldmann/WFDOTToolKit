using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;

namespace WFDOTToolKit
{
    public class WFDOTDeploy
    {
        public static void ExcelToUnicodeText()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            var res = dialog.ShowDialog();
            if (res != DialogResult.OK) return;
            var deployPath = Path.Combine(Path.GetDirectoryName(dialog.FileName), "deploy.txt");
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(dialog.FileName);
            wb.SaveAs(deployPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlUnicodeText);
            wb.Close(false);
            app.Quit();
        }
    }
}
