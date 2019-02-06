using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestFramework
{
    class Program
    {
        private static Excel._Worksheet Worksheet = null;
        private static object locker = new object();
        private static StreamWriter stream = null;
        #region Example
        static int x = 0;
        static void Main1(string[] args)
        {
            for (int i = 0; i < 5; i++)
            {
                Thread myThread = new Thread(Count);
                myThread.Name = "Поток " + i.ToString();
                myThread.Start();
            }

            Console.ReadLine();
        }
        public static void Count()
        {
            lock (locker)
            {
                x = 1;
                for (int i = 1; i < 9; i++)
                {
                    Console.WriteLine("{0}: {1}", Thread.CurrentThread.Name, x);
                    x++;
                    Thread.Sleep(100);
                }
            }
        }
        #endregion

        static void Main(string[] args)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Open(@"C:\Users\danon\Desktop\qwerty\example");
            //Excel._Worksheet workSheet = excelApp.ActiveSheet; for sync way
            Worksheet = excelApp.ActiveSheet;
            using (var writer = File.AppendText(@"C:\Users\danon\Desktop\qwerty\output.txt"))
            {
                stream = writer;
                // synchronous way
                ProcessAndWriteCells(new Range(2, 10000));
            }
            // asynchronous way
            //WriteFileFunctionWithThreads(10000);

            stopwatch.Stop();
            Console.WriteLine("Time elapsed (in ms) : " + stopwatch.ElapsedMilliseconds);
            Console.WriteLine("Ready to go");
            Console.ReadLine();
        }

        /*private static void WriteFileFunctionWithThreads(int cellsCount)
        {
            var countStep = cellsCount / 4;
            Thread[] additionalThreads = {
                new Thread(new ParameterizedThreadStart(Process)),
                new Thread(new ParameterizedThreadStart(Process)),
                new Thread(new ParameterizedThreadStart(Process))
            };
            var currentStep = countStep;
            for (var i = 0; i < additionalThreads.Length; i++)
            {
                additionalThreads[i].Start(new Range(currentStep, currentStep + countStep, i+2));
                currentStep += countStep;
            }
            
            ProcessAndWriteCells(new Range(2, countStep));
        }

        private static void Process(object obj)
        {
            Range range = null;
            //try
            //{
                range = (Range)obj;

                var check = range.StartingCellIndex;
            //}
            //catch (Exception e)
            //{
            //Console.WriteLine("Unexpected range\r\n");
            //Console.WriteLine(e.Message);
            // return;
            //}
            ProcessAndWriteCells(range);
        }*/

        private static void ProcessAndWriteCells(Range range)
        { 
            for (var i = 2; i < 10001; i++)
            {
                string value = Convert.ToString(Worksheet.get_Range($"B{i}", Type.Missing).Value2);
                var address = ProcessStringAddress(value);

                if (address == string.Empty)
                    continue;

                //lock (locker)
                //{
                    
                        stream.WriteLine(Worksheet.get_Range($"A{i}", Type.Missing).Value2 + ";" + address);
                    
               // }

                //Console.WriteLine($"Value processed by thread : {range.HandlingThreadNumber}");
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