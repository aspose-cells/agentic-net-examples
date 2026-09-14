// Title: Measure and compare heap memory usage of Aspose.Cells default and custom calculation engines on a large workbook (C#)
// AI Prompts: Generate C# code that builds a 5,000‑row workbook with numeric data and a SUM formula per row, then uses GC.GetTotalMemory to profile the heap consumption of the built‑in calculation engine. | Write a C# example that implements a no‑op AbstractCalculationEngine, runs Workbook.CalculateFormula with CalculationOptions.CustomEngine, and outputs the memory delta versus the default engine. | Provide a C# snippet that records both execution time (using Stopwatch) and heap memory before and after calculation for default and custom engines, and logs the results to the console.
// Common Searches: Aspose.Cells how to profile memory usage of calculation engine in .NET | compare heap allocation of default vs custom AbstractCalculationEngine in C# | benchmark formula calculation memory consumption for large Excel workbook using Aspose.Cells | measure GC.GetTotalMemory before and after Workbook.CalculateFormula Aspose.Cells
// Tags: Aspose.Cells calculation engine heap profiling | no‑op custom AbstractCalculationEngine benchmark | large Excel workbook formula memory impact | GC.GetTotalMemory measurement Aspose.Cells | default calculation engine memory usage comparison

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryProfiling
{
    // Simple custom calculation engine that does not process any functions.
    // The default engine will handle all calculations.
    // // Creates a 5,000‑row workbook with numeric values and row‑wise SUM formulas, then measures the heap memory consumed during formula calculation using the built‑in engine and a no‑op custom AbstractCalculationEngine, reporting the memory difference.
    public class NoOpCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // No custom processing; let the default engine handle the function.
            // By not setting ProcessBuiltInFunctions and not handling any function,
            // the engine effectively delegates all work to the built‑in engine.
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a large workbook with many formulas to stress the calculation engine.
            Workbook wbDefault = CreateLargeWorkbook();

            // Measure memory usage with the default calculation engine.
            long memoryDefault = MeasureCalculationMemory(wbDefault, null);
            Console.WriteLine($"Memory used with default engine: {memoryDefault:N0} bytes");

            // Create another workbook (identical data) for custom engine measurement.
            Workbook wbCustom = CreateLargeWorkbook();

            // Measure memory usage with a custom calculation engine.
            var customEngine = new NoOpCustomEngine();
            var options = new CalculationOptions { CustomEngine = customEngine };
            long memoryCustom = MeasureCalculationMemory(wbCustom, options);
            Console.WriteLine($"Memory used with custom engine: {memoryCustom:N0} bytes");

            // Show the difference.
            long diff = memoryCustom - memoryDefault;
            Console.WriteLine($"Difference (custom - default): {diff:N0} bytes");
        }

        // Creates a workbook with a sizable number of rows and columns filled with formulas.
        static Workbook CreateLargeWorkbook()
        {
            // Create a new workbook.
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            const int rows = 5000;   // Adjust for desired size.
            const int cols = 10;

            // Populate numeric data.
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(r + c);
                }
            }

            // Add a formula in the last column that sums the row.
            for (int r = 0; r < rows; r++)
            {
                // Example: =SUM(A1:J1) for each row.
                string formula = $"=SUM(A{r + 1}:{GetColumnName(cols - 1)}{r + 1})";
                cells[r, cols].Formula = formula;
            }

            return wb;
        }

        // Measures the heap memory consumed during formula calculation.
        // If calcOptions is null, the default engine is used.
        static long MeasureCalculationMemory(Workbook wb, CalculationOptions calcOptions)
        {
            // Force a full garbage collection before measurement.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record memory before calculation.
            long before = GC.GetTotalMemory(true);

            // Perform calculation.
            if (calcOptions == null)
                wb.CalculateFormula();               // Default engine.
            else
                wb.CalculateFormula(calcOptions);    // Custom engine.

            // Force another GC to clean up temporary objects created during calculation.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record memory after calculation.
            long after = GC.GetTotalMemory(true);

            // Return the difference.
            return after - before;
        }

        // Helper to convert a zero‑based column index to Excel column letters (A, B, …).
        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string name = "";
            do
            {
                name = letters[index % 26] + name;
                index = index / 26 - 1;
            } while (index >= 0);
            return name;
        }
    }
}
