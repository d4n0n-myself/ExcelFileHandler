using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            //Example();
            //MyImplementation();
            WriteFileFunction();
        }


        private static void WriteFileFunction()
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Open(@"C:\Users\danon\Desktop\qwerty\example");
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            using (var stream = File.AppendText(@"C:\Users\danon\Desktop\qwerty\output.txt"))
            {
                for (var i = 2; i < 957384; i++)
                {
                    string value = Convert.ToString(workSheet.get_Range($"B{i}", Type.Missing).Value2);
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
                        continue;
                    index = index + 2;

                    stream.Write(workSheet.get_Range($"A{i}", Type.Missing).Value2 + ";" + 
                        value.Substring(index).Replace("д.", "").Replace("ул.", "").Replace("кв.", "") + "\r\n");
                }
            }
            Console.WriteLine("Ready to go");
            Console.ReadLine();
        }

        private static void MyImplementation()
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Open(@"C:\Users\danon\Desktop\qwerty\1234");
            Excel._Worksheet workSheet = excelApp.ActiveSheet;
            List<dynamic> completeValues = new List<dynamic>();
            string a;
            for (var i = 1; i < 16; i++)
            {
                string value = Convert.ToString(workSheet.get_Range($"A{i}", Type.Missing).Value2);
                File.AppendAllText(@"C:\Users\danon\Desktop\qwerty\output.txt", workSheet.get_Range($"A{i}", Type.Missing).Value2 + "\r\n");
            }

            Console.WriteLine("Ready to go");
            Console.ReadLine();
        }

        static void Example()
        {
            var excelApp = new Excel.Application();

            excelApp.Visible = true;

            //excelApp.Workbooks.Add();



            excelApp.Workbooks.Open(@"C:\Users\danon\Desktop\qwerty\1234");
            Excel._Worksheet workSheet = excelApp.ActiveSheet;
            workSheet.Cells[1, "A"] = "ID Number";

            string a;
            var excelcells = workSheet.get_Range("A1", Type.Missing);
            a = Convert.ToString(excelcells.Value2);

            Console.WriteLine("a1=" + a);


            Console.ReadLine();
        }
    }
}