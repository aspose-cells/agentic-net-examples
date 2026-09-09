// Title: Compare Workbook.CalculateFormula with per‑cell Cell.Calculate performance on a 10,000‑row workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that builds a workbook with 10,000 rows of formulas, clones it, and records the elapsed time for Workbook.CalculateFormula and for Cell.Calculate on each formula cell. | Create code that gathers all formula cells, applies CalculationOptions, and logs the duration of direct per‑cell evaluation using Aspose.Cells. | Handle any exceptions and output the milliseconds taken by both the calculation chain and the individual cell calculations.
// Common Searches: Aspose.Cells benchmark Workbook.CalculateFormula vs Cell.Calculate for large worksheets | How to benchmark formula calculation time in Aspose.Cells .NET | C# example comparing calculation chain and per‑cell evaluation timing in Aspose.Cells | Benchmarking formula evaluation on a 10k‑row workbook with Aspose.Cells
// Tags: Aspose.Cells formula calculation timing | Workbook.CalculateFormula performance analysis | Cell.Calculate execution profiling | large worksheet evaluation .NET | calculation chain versus direct evaluation | Aspose.Cells performance profiling

using System;
using System.Diagnostics;
using System.Linq;
using Aspose.Cells;

// The program creates a workbook containing 10,000 rows of simple formulas, clones it for two tests, then measures and prints the milliseconds taken by the built‑in calculation chain (Workbook.CalculateFormula) and by evaluating each formula cell individually with Cell.Calculate.
class BenchmarkFormulaCalculation
{
    static void Main()
    {
        try
        {
            // Create a workbook with a large set of formulas
            Workbook wbOriginal = new Workbook();
            Worksheet ws = wbOriginal.Worksheets[0];

            // Populate column A with numbers 1..10000
            int rowCount = 10000;
            for (int i = 0; i < rowCount; i++)
            {
                ws.Cells[i, 0].PutValue(i + 1); // A column
            }

            // Add formulas in column B: =A*2
            for (int i = 0; i < rowCount; i++)
            {
                ws.Cells[i, 1].Formula = $"=A{i + 1}*2";
            }

            // Add formulas in column C: =B+A
            for (int i = 0; i < rowCount; i++)
            {
                ws.Cells[i, 2].Formula = $"=B{i + 1}+A{i + 1}";
            }

            // Clone the workbook for separate tests
            Workbook wbForChain = new Workbook();
            wbForChain.Copy(wbOriginal);

            Workbook wbForDirect = new Workbook();
            wbForDirect.Copy(wbOriginal);

            // -------------------------
            // Benchmark using calculation chain (Workbook.CalculateFormula)
            // -------------------------
            Stopwatch swChain = Stopwatch.StartNew();
            wbForChain.CalculateFormula(); // uses internal calculation chain
            swChain.Stop();

            // -------------------------
            // Benchmark using direct evaluation (Cell.Calculate for each formula cell)
            // -------------------------
            // Find all cells that contain formulas
            var formulaCells = wbForDirect.Worksheets[0].Cells
                .Cast<Cell>()
                .Where(c => c.IsFormula)
                .ToList();

            CalculationOptions calcOptions = new CalculationOptions(); // default options

            Stopwatch swDirect = Stopwatch.StartNew();
            foreach (Cell cell in formulaCells)
            {
                // Directly evaluate each cell's formula
                cell.Calculate(calcOptions);
            }
            swDirect.Stop();

            // Output the results
            Console.WriteLine($"Calculation chain time: {swChain.ElapsedMilliseconds} ms");
            Console.WriteLine($"Direct evaluation time: {swDirect.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
