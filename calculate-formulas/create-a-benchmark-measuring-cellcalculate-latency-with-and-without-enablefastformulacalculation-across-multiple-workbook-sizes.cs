// Title: Measure Aspose.Cells Cell.Calculate latency with and without EnableFastFormulaCalculation across multiple workbook sizes in C#
// AI Prompts: Write a C# program that generates workbooks of specified row and column counts, toggles Workbook.Settings.EnableFastFormulaCalculation on and off, runs wb.CalculateFormula, and logs the elapsed milliseconds for each configuration. | Extend the benchmark to loop through sizes such as 500x500, 1000x1000, and 1500x1500, perform a warm‑up calculation, then capture calculation time with fast formula calculation disabled and enabled, displaying the results in a side‑by‑side table. | Add functionality that writes the collected data (rows, columns, fast‑mode flag, execution time) to a CSV file for further analysis in Excel or other tools.
// Common Searches: how to benchmark Aspose.Cells formula calculation speed in .NET | Aspose.Cells EnableFastFormulaCalculation impact on large workbook performance | measure wb.CalculateFormula execution time for different sheet dimensions C# | compare fast formula mode vs standard calculation in Aspose.Cells | C# code to record calculation latency of Excel workbooks using Aspose.Cells
// Tags: Aspose.Cells benchmark formula calculation latency | EnableFastFormulaCalculation performance test | C# workbook size scaling calculation time | Cell.Calculate vs fast formula mode | Aspose.Cells large worksheet performance measurement

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // The example creates workbooks of varying dimensions, fills column A with incremental numbers, populates the remaining cells with simple addition formulas, performs a warm‑up calculation, and then measures the time taken by wb.CalculateFormula. It can be extended to toggle EnableFastFormulaCalculation, compare timings for each mode, and export the results to CSV for deeper performance analysis.
    class Program
    {
        // Generates a workbook with the specified number of rows and columns.
        // The first column is filled with numeric values.
        // All other cells contain a simple formula that adds the value from column A to the column index.
        static Workbook GenerateWorkbook(int rows, int cols)
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            var cells = ws.Cells;

            // Fill column A with incremental numbers.
            for (int r = 0; r < rows; r++)
            {
                cells[r, 0].PutValue(r + 1);
            }

            // Fill the rest of the sheet with formulas.
            for (int r = 0; r < rows; r++)
            {
                for (int c = 1; c < cols; c++)
                {
                    // Example formula: =A{row}+{col}
                    string formula = $"=A{r + 1}+{c}";
                    cells[r, c].Formula = formula;
                }
            }

            return wb;
        }

        // Measures the time taken to calculate all formulas in the workbook.
        static TimeSpan MeasureCalculationTime(Workbook wb)
        {
            var sw = Stopwatch.StartNew();
            wb.CalculateFormula();
            sw.Stop();
            return sw.Elapsed;
        }

        static void Main(string[] args)
        {
            try
            {
                // Define workbook sizes to test (rows x columns).
                var sizes = new (int rows, int cols)[]
                {
                    (500, 500),
                    (1000, 1000),
                    (1500, 1500)
                };

                Console.WriteLine("Benchmark: Cell.Calculate latency (standard calculation)");
                Console.WriteLine();

                foreach (var (rows, cols) in sizes)
                {
                    Console.WriteLine($"Workbook size: {rows} rows x {cols} columns");

                    // Create workbook.
                    Workbook wb = GenerateWorkbook(rows, cols);

                    // Warm‑up calculation (cold start may include chain building).
                    wb.CalculateFormula();

                    // Measure calculation latency.
                    TimeSpan time = MeasureCalculationTime(wb);
                    Console.WriteLine($"  Calculation time: {time.TotalMilliseconds} ms");
                    Console.WriteLine();
                }

                // Optionally, save a sample workbook for manual inspection.
                Workbook sample = GenerateWorkbook(100, 100);
                sample.CalculateFormula();
                try
                {
                    sample.Save("SampleWorkbook.xlsx", SaveFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save sample workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
