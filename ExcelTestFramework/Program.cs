using System;
using System.Diagnostics;

namespace ExcelTestFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            ExcelService.WriteFileFunction();
            Helpers.SplitFileInTwoParts();
            
            stopwatch.Stop();
            Console.WriteLine("Time elapsed (in ms) : " + stopwatch.ElapsedMilliseconds);
            Console.WriteLine("Ready to go");
            Console.ReadLine();
        }
    }
}