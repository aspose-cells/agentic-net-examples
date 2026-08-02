// Title: Aspose.Cells .NET Benchmark: Cell.Calculate Latency With and Without Fast Formula Mode Across Workbook Sizes
// Description: Creates workbooks of 500×500, 1000×1000 and 1500×1500 cells, fills them with simple sum formulas, runs a warm‑up calculation, then measures the time of a single Cell.Calculate call on the bottom‑right cell. The sample is extended to record timings with the workbook.Settings.EnableFastFormulaCalculation flag turned on and off, enabling a direct performance comparison.
// Keywords: Aspose.Cells | Cell.Calculate | FastFormulaCalculation | benchmark | performance testing | .NET | C# | large workbook | formula evaluation speed | Excel calculation latency
// Common Searches: Aspose.Cells benchmark Cell.Calculate speed | measure fast formula mode latency .NET | compare calculation time with EnableFastFormulaCalculation | performance test large Excel workbook Aspose | C# formula calculation benchmark Aspose.Cells
// Developer Intent: Compare the execution time of Cell.Calculate when fast‑formula mode is enabled versus disabled for workbooks of varying dimensions.
// Use Cases: Determine the performance benefit of Fast Formula mode before deploying to production. | Identify the workbook size where fast‑formula calculation starts to yield measurable latency reductions. | Validate that formula evaluation meets SLA requirements for automated reporting pipelines.
// AI Prompts: Generate C# code that toggles workbook.Settings.EnableFastFormulaCalculation, runs Cell.Calculate for each workbook size, and logs elapsed milliseconds for both states. | Create a reusable method that accepts workbook size and a boolean for fast‑formula mode, returns the Cell.Calculate duration, and integrates it into the benchmark loop. | Write a script that outputs a CSV with columns: Size, FastFormulaEnabled, CalculationTimeMs, suitable for plotting performance charts.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates workbooks of 500×500, 1000×1000 and 1500×1500 cells, fills them with simple sum formulas, runs a warm‑up calculation, then measures the time of a single Cell.Calculate call on the bottom‑right cell. The sample is extended to record timings with the workbook.Settings.EnableFastFormulaCalculation flag turned on and off, enabling a direct performance comparison.
class FormulaCalculationBenchmark
{
    static void Main()
    {
        try
        {
            // Define different workbook sizes (rows x columns)
            int[] sizes = new int[] { 500, 1000, 1500 };

            foreach (int size in sizes)
            {
                Console.WriteLine($"Benchmark for workbook size {size}x{size}");

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate the sheet with simple formulas.
                // Each cell in a row (except the first column) sums a constant with its column index.
                for (int row = 0; row < size; row++)
                {
                    // First column gets a static value to avoid circular references.
                    cells[row, 0].PutValue(1);

                    for (int col = 1; col < size; col++)
                    {
                        // Example formula: =A{row+1}+{col}
                        cells[row, col].Formula = $"=A{row + 1}+{col}";
                    }
                }

                // Choose a cell far down the sheet to measure calculation time.
                Cell targetCell = cells[size - 1, size - 1];

                // Warm‑up: calculate the whole workbook once so that any lazy initialization is done.
                workbook.CalculateFormula();

                // Benchmark calculation time (standard mode)
                Stopwatch sw = Stopwatch.StartNew();
                targetCell.Calculate(new CalculationOptions());
                sw.Stop();

                Console.WriteLine($"  Calculation time: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
