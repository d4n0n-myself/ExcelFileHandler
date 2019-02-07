using System.IO;
using System.Linq;

namespace ExcelTestFramework
{
    internal static class Helpers
    {
        internal static void SplitFileInTwoParts()
        {
            var str = File.ReadLines("ready-output.txt");
            File.WriteAllLines("ready-output1.txt", str.Take(500000));
            File.WriteAllLines("ready-output2.txt", str.Skip(500000));
        }
    }
}
