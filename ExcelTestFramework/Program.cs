using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            //WriteFileFunction();
            //ImproveTxtFile();
            //Cut();
            Skip();
            
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

        private static void ImproveTxtFile()
        {
            var str = File.ReadLines("output.txt");
            var processedStrings = new List<string>();

            foreach (var s in str)
            {
                var origArray = s.Split(';');
                var processedString = origArray[1];
                processedString = processedString.Replace("пр-кт","").Replace("б-р.","");
                var index = processedString.IndexOf("п.");
                if (index != -1)
                {
                    index += 2;
                    processedString = processedString.Substring(index);
                }

                var temp = origArray[0] + ';' + processedString;
                processedStrings.Add(temp.Replace("  "," ").Replace(", ",",").Replace(" ,",","));
            }
            
            File.WriteAllLines("ready-output.txt", processedStrings);
        }

        private static void Cut()
        {
            var str = File.ReadLines("ready-output.txt");
            File.WriteAllLines("ready-output1.txt", str.Take(500000));
        }

        private static void Skip()
        {
            var str = File.ReadLines("ready-output.txt");
            File.WriteAllLines("ready-output2.txt", str.Skip(500000));
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
                           value.IndexOf("п.", 2) : value.IndexOf("с.")
                   : value.IndexOf("г.");
            }

            if (index <= 0)
                return "";
            index = index + 2;

            return value.Substring(index).Replace("д.", "").Replace("ул.", "").Replace("кв.", "");
        }
    }
}