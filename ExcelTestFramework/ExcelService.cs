using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestFramework
{
    internal class ExcelService
    {
        internal static void WriteFileFunction()
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Open(@"C:\Users\danon\Desktop\qwerty\example");
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            using (var stream = File.AppendText(@"C:\Users\danon\Desktop\qwerty\newoutput.txt"))
            {
                for (var i = 2; i < 957384; i++)
                {
                    string value = Convert.ToString(workSheet.get_Range($"B{i}", Type.Missing).Value2);
                    var address = ProcessStringAddress(value);

                    if (address == string.Empty)
                        continue;

                    stream.WriteLine(workSheet.get_Range($"A{i}", Type.Missing).Value2 + ";" + address);
                }
            }
        }

        private static string ProcessStringAddress(string value)
        {
            value = value.Replace("д.", "")
                .Replace("ул.", "")
                .Replace("кв.", "")
                .Replace("ш.", "")
                .Replace("Россия, респ.Татарстан,", "")
                .Replace("пр-кт", "")
                .Replace("б-р.", "")
                .Replace("п.", "");

            int index;

            if (value.IndexOf("д.") != -1 && !Char.IsNumber(value[value.IndexOf("д.") + 2]))
                index = value.IndexOf("д.");
            else if (value.IndexOf("пгт.") != -1)
                index = value.IndexOf("пгт.") + 2;
            else
            {
                index = value.IndexOf("г.") == -1 ?
                       value.IndexOf("с.") == -1 ?
                           value.IndexOf("п.", 0, 2) : value.IndexOf("с.")
                   : value.IndexOf("г.");
            }

            if (index <= 0)
                return "";
            index = index + 2;

            return value.Substring(index).Replace("  ", " ").Replace(", ", ",").Replace(" ,", ",");
        }
    }
}
