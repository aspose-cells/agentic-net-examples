// Title: Measure and compare Aspose.Cells CalculateFormula execution time with and without cached values in C#
// Common Searches: how to measure Aspose.Cells CalculateFormula execution time in C# | benchmark Excel formula calculation performance using Aspose.Cells .NET | compare first and second CalculateFormula call timings in Aspose.Cells | use Stopwatch to profile workbook calculation speed with Aspose.Cells | performance impact of recalculating a workbook after previous calculation in Aspose.Cells
// Tags: Aspose.Cells CalculateFormula execution time measurement | Excel workbook calculation performance .NET | profiling Aspose.Cells formula recalculation | stopwatch based timing of workbook calculations | impact of cached values on Aspose.Cells calculation speed

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsCalculationChainDemo
{
    // The sample creates a workbook with 5,000 rows of random numbers and a sum formula per row, then uses Stopwatch to record the time of an initial full CalculateFormula call, invokes CalculateFormula again to force a fresh recalculation, records the second duration, prints both timings, and saves the file as CalculationChainDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                worksheet.Name = "DataSheet";

                // Populate the worksheet with data and formulas
                const int rows = 5000;
                const int cols = 20;

                var rnd = new Random();
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        worksheet.Cells[i, j].PutValue(rnd.NextDouble() * 100);
                    }

                    // Add a formula that sums the row
                    string sumFormula = $"=SUM(A{i + 1}:{CellIndexToName(cols - 1)}{i + 1})";
                    worksheet.Cells[i, cols].Formula = sumFormula;
                }

                // -------------------------------------------------
                // 1. Full calculation (default behavior)
                // -------------------------------------------------
                var swFullCalc = Stopwatch.StartNew();
                workbook.CalculateFormula(); // Calculates the whole workbook
                swFullCalc.Stop();

                // -------------------------------------------------
                // 2. Re‑calculation after clearing cached values
                // -------------------------------------------------
                // Clear previously calculated values to force a fresh calculation
                workbook.CalculateFormula(); // Re‑calculate again for timing comparison
                var swRecalc = Stopwatch.StartNew();
                workbook.CalculateFormula();
                swRecalc.Stop();

                // -------------------------------------------------
                // Output the timing results
                // -------------------------------------------------
                Console.WriteLine($"Initial full calculation time: {swFullCalc.ElapsedMilliseconds} ms");
                Console.WriteLine($"Re‑calculation time: {swRecalc.ElapsedMilliseconds} ms");

                // Save the workbook (no need to check existence before saving)
                workbook.Save("CalculationChainDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to convert a zero‑based column index to Excel column name (A, B, ..., Z, AA, AB, ...)
        private static string CellIndexToName(int columnIndex)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string name = string.Empty;
            int dividend = columnIndex + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                name = letters[modulo] + name;
                dividend = (dividend - modulo) / 26;
            }

            return name;
        }
    }
}
