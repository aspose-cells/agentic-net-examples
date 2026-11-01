// Title: C# benchmark: comparing cell.Calculate vs Workbook.CalculateFormula for a 20,000‑row workbook in Aspose.Cells
// Common Searches: how to measure Aspose.Cells formula calculation performance in C# | benchmark cell.Calculate versus Workbook.CalculateFormula in Aspose.Cells | performance test for evaluating 20k rows of formulas with Aspose.Cells | compare individual cell calculation and full workbook calculation speed Aspose.Cells .NET | timing Aspose.Cells formula engine for large worksheets in C#
// Tags: cell.Calculate performance test | Workbook.CalculateFormula speed test | Aspose.Cells large workbook formula benchmark | C# formula calculation engine comparison | individual cell evaluation vs calculation chain

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a 20,000‑row workbook, populates numeric values and two formula columns, clears cached results, then measures execution time for two approaches: iterating each formula cell with cell.Calculate(new CalculationOptions()) and invoking Workbook.CalculateFormula() for the whole sheet, outputting the elapsed milliseconds for each method.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and fill it with data and formulas
            Workbook original = new Workbook();
            Worksheet ws = original.Worksheets[0];
            int rowCount = 20000; // number of rows for the benchmark

            // Column A: simple numeric values
            for (int i = 0; i < rowCount; i++)
            {
                ws.Cells[i, 0].PutValue(i + 1);
            }

            // Column B: simple formula =A1*2
            for (int i = 0; i < rowCount; i++)
            {
                ws.Cells[i, 1].Formula = $"A{i + 1}*2";
            }

            // Column C: more complex formula =B1+A1
            for (int i = 0; i < rowCount; i++)
            {
                ws.Cells[i, 2].Formula = $"B{i + 1}+A{i + 1}";
            }

            // Ensure the workbook is in a known state (no cached results)
            original.CalculateFormula();

            // -------------------- Direct evaluation (cell.Calculate) --------------------
            Workbook wbDirect = new Workbook();
            wbDirect.Copy(original); // deep copy of the original workbook
            Worksheet wsDirect = wbDirect.Worksheets[0];
            Stopwatch swDirect = Stopwatch.StartNew();

            foreach (Cell cell in wsDirect.Cells)
            {
                if (cell.IsFormula)
                {
                    // Evaluate each formula individually using default calculation options
                    cell.Calculate(new CalculationOptions());
                }
            }

            swDirect.Stop();

            // -------------------- Calculation chain (Workbook.CalculateFormula) --------------------
            Workbook wbChain = new Workbook();
            wbChain.Copy(original); // deep copy of the original workbook
            Stopwatch swChain = Stopwatch.StartNew();

            wbChain.CalculateFormula(); // evaluate the whole workbook using the calculation engine

            swChain.Stop();

            // Output the benchmark results
            Console.WriteLine($"Direct evaluation time: {swDirect.ElapsedMilliseconds} ms");
            Console.WriteLine($"Calculation chain time: {swChain.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
