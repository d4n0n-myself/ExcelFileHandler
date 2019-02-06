namespace ExcelTestFramework
{
    class Range
    {
        public Range(int start, int end, int threadNumber = 1)
        {
            StartingCellIndex = start;
            EndingCellIndex = end;
            HandlingThreadNumber = threadNumber;
        }

        public int StartingCellIndex { get; }
        public int EndingCellIndex { get; }

        // unnecessary property, purpose: more information in console
        public int HandlingThreadNumber { get; } 
    }
}
