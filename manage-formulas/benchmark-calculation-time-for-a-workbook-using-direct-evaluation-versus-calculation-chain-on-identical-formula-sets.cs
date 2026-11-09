// Title: Compare Aspose.Cells calculation chain vs per‑cell evaluation performance for a 1000‑row workbook in C#
// Common Searches: Aspose.Cells how to measure calculation time of CalculateFormula versus Cell.Calculate in C# | benchmark formula recalculation speed for large worksheets using Aspose.Cells .NET | performance test for Aspose.Cells calculation chain on 1000 rows C# example | compare per‑cell formula evaluation and workbook calculation chain in Aspose.Cells | C# code to time Aspose.Cells formula evaluation methods
// Tags: Aspose.Cells calculation chain performance | Cell.Calculate direct evaluation timing | benchmark workbook formula recalculation .NET | large worksheet formula performance Aspose.Cells | C# Aspose.Cells formula evaluation benchmark

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace Example
{
    // The example creates a 1,000‑row workbook where column C contains formulas referencing columns A and B, builds the calculation chain once, then measures and compares the elapsed time of a full workbook recalculation using CalculateFormula() against individual cell calculations using Cell.Calculate(), outputting both timings.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the sheet with data and formulas
                // 1000 rows, column C contains a formula referencing columns A and B
                for (int row = 0; row < 1000; row++)
                {
                    sheet.Cells[row, 0].PutValue(row);                     // Column A
                    sheet.Cells[row, 1].PutValue(row * 2);                 // Column B
                    sheet.Cells[row, 2].Formula = $"=A{row + 1}+B{row + 1}"; // Column C formula (1‑based Excel rows)
                }

                // Build the calculation chain once before benchmarking
                workbook.CalculateFormula();

                // Benchmark the calculation chain (recalculates the whole workbook)
                Stopwatch swChain = Stopwatch.StartNew();
                workbook.CalculateFormula(); // Uses the internal calculation chain
                swChain.Stop();

                // Benchmark direct evaluation (calculates each cell individually)
                CalculationOptions calcOptions = new CalculationOptions();
                Stopwatch swDirect = Stopwatch.StartNew();
                for (int row = 0; row < 1000; row++)
                {
                    // Calculate the formula in the current cell directly
                    sheet.Cells[row, 2].Calculate(calcOptions);
                }
                swDirect.Stop();

                // Output the timing results
                Console.WriteLine($"Calculation chain time: {swChain.ElapsedMilliseconds} ms");
                Console.WriteLine($"Direct evaluation time: {swDirect.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
