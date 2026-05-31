using System;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    class Program
    {
        static void Main()
        {
            // Define size of the test data set
            const int rows = 5000;   // number of rows
            const int cols = 10;     // number of columns with formulas

            // -------------------- Scenario 1: Calculation chain disabled --------------------
            Workbook wbNoChain = new Workbook();                     // create workbook
            FillData(wbNoChain, rows, cols);                        // populate with formulas
            wbNoChain.Settings.FormulaSettings.EnableCalculationChain = false; // disable chain

            // Force garbage collection before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long memBeforeNoChain = GC.GetTotalMemory(true);        // memory before calculation
            wbNoChain.CalculateFormula();                           // calculate formulas
            long memAfterNoChain = GC.GetTotalMemory(true);         // memory after calculation

            Console.WriteLine("=== Calculation Chain Disabled ===");
            Console.WriteLine($"Memory before calculation: {FormatBytes(memBeforeNoChain)}");
            Console.WriteLine($"Memory after  calculation: {FormatBytes(memAfterNoChain)}");
            Console.WriteLine($"Memory increase: {FormatBytes(memAfterNoChain - memBeforeNoChain)}");

            // Save the workbook (uses the provided save rule)
            wbNoChain.Save("LargeDataset_NoChain.xlsx");

            // -------------------- Scenario 2: Calculation chain enabled --------------------
            Workbook wbWithChain = new Workbook();                   // create another workbook
            FillData(wbWithChain, rows, cols);                      // populate with the same data
            wbWithChain.Settings.FormulaSettings.EnableCalculationChain = true; // enable chain

            // Force garbage collection before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long memBeforeWithChain = GC.GetTotalMemory(true);      // memory before calculation
            wbWithChain.CalculateFormula();                         // calculate formulas
            long memAfterWithChain = GC.GetTotalMemory(true);       // memory after calculation

            Console.WriteLine("\n=== Calculation Chain Enabled ===");
            Console.WriteLine($"Memory before calculation: {FormatBytes(memBeforeWithChain)}");
            Console.WriteLine($"Memory after  calculation: {FormatBytes(memAfterWithChain)}");
            Console.WriteLine($"Memory increase: {FormatBytes(memAfterWithChain - memBeforeWithChain)}");

            // Save the workbook (uses the provided save rule)
            wbWithChain.Save("LargeDataset_WithChain.xlsx");
        }

        // Populates the workbook with a large set of inter‑dependent formulas.
        // Each cell in column B..K contains a formula that adds the value of the cell
        // to the left plus a constant, creating a long dependency chain.
        static void FillData(Workbook workbook, int rows, int cols)
        {
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Initialize first column with numeric values
            for (int r = 0; r < rows; r++)
            {
                cells[r, 0].PutValue(r + 1); // A column
            }

            // Create formulas that depend on the previous column
            for (int c = 1; c < cols; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    // Example formula: =B1+1 (where B is the previous column)
                    string colLetterPrev = CellsHelper.ColumnIndexToName(c - 1);
                    string colLetterCurr = CellsHelper.ColumnIndexToName(c);
                    cells[r, c].Formula = $"={colLetterPrev}{r + 1}+1";
                }
            }
        }

        // Helper to format byte values into a readable string.
        static string FormatBytes(long bytes)
        {
            const long KB = 1024;
            const long MB = KB * 1024;
            const long GB = MB * 1024;

            if (bytes >= GB) return $"{bytes / (double)GB:F2} GB";
            if (bytes >= MB) return $"{bytes / (double)MB:F2} MB";
            if (bytes >= KB) return $"{bytes / (double)KB:F2} KB";
            return $"{bytes} B";
        }
    }
}