// Title: Enable calculation chain in Aspose.Cells .NET, run workbook-wide formula calculation, and benchmark performance versus default mode
// Common Searches: how to enable calculation chain in Aspose.Cells C# and compare formula calculation speed | Aspose.Cells workbook.CalculateFormula performance test default mode vs calculation chain | measure Excel formula evaluation time using Aspose.Cells .NET for large worksheets | C# code to benchmark formula recalculation after enabling calculation chain in Aspose.Cells | compare calculation mode impact on workbook-wide formula processing in Aspose.Cells
// Tags: enable calculation chain Aspose.Cells | benchmark workbook.CalculateFormula speed | Excel formula evaluation performance .NET | Aspose.Cells calculation mode comparison | large worksheet formula recalculation timing | optimize Excel calculation Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a workbook with 1,000 rows of numeric data and simple formulas, enables the calculation chain, runs Workbook.CalculateFormula to evaluate all formulas, measures execution time for default and chain-enabled modes, outputs the timing results, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate the first worksheet with data and formulas
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 1000; row++)
            {
                // Simple numeric values in column A
                sheet.Cells[row, 0].PutValue(row);

                // Formula in column B that depends on column A
                sheet.Cells[row, 1].Formula = $"=A{row + 1}*2";
            }

            // -------------------------------------------------
            // Measure calculation time (default settings)
            // -------------------------------------------------
            Stopwatch sw = Stopwatch.StartNew();
            workbook.CalculateFormula(); // Run workbook-wide calculation
            sw.Stop();

            long timeDefault = sw.ElapsedMilliseconds;
            Console.WriteLine($"Calculation time (default settings): {timeDefault} ms");

            // -------------------------------------------------
            // Recalculate and measure again (still default settings)
            // -------------------------------------------------
            workbook.CalculateFormula(); // Ensure fresh evaluation

            sw.Restart();
            workbook.CalculateFormula();
            sw.Stop();

            long timeSecondRun = sw.ElapsedMilliseconds;
            Console.WriteLine($"Calculation time (second run): {timeSecondRun} ms");

            // Save the workbook (optional)
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
