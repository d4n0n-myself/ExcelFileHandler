using System;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            WriteFileFunction();

            stopwatch.Stop();
            Console.WriteLine("Time elapsed (in ms) : " + stopwatch.ElapsedMilliseconds);
            Console.WriteLine("Ready to go");
            Console.ReadLine();
        }

        private static void WriteFileFunction()
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Open(@"C:\Users\danon\Desktop\qwerty\example");
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            using (var stream = File.AppendText(@"C:\Users\danon\Desktop\qwerty\output.txt"))
            {
                for (var i = 2; i < 10000; i++)
                {
                    string value = Convert.ToString(workSheet.get_Range($"B{i}", Type.Missing).Value2);
                    var address = ProcessStringAddress(value);

                    if (address == string.Empty)
                        continue;

                    stream.Write(workSheet.get_Range($"A{i}", Type.Missing).Value2 + ";" + address + "\r\n");
                }
            }
        }

        private static string ProcessStringAddress(string value)
        {
            int index;
            if (value.IndexOf("д.") != -1 && !Char.IsNumber(value[value.IndexOf("д.") + 2]))
                index = value.IndexOf("д.");
            else if (value.IndexOf("пгт.") != -1)
                index = value.IndexOf("пгт.") + 2;
            else
            {
                index = value.IndexOf("г.") == -1 ?
                       value.IndexOf("с.") == -1 ?
                           value.IndexOf("п.") : value.IndexOf("с.")
                   : value.IndexOf("г.");
            }

            if (index <= 0)
                return "";
            index = index + 2;

            return value.Substring(index).Replace("д.", "").Replace("ул.", "").Replace("кв.", "");
        }
    }
}