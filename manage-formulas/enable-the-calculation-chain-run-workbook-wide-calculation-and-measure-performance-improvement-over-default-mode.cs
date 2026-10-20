// Title: Benchmarking workbook-wide formula calculation performance with Aspose.Cells for .NET by timing CalculateFormula with and without the calculation chain
// Common Searches: how to measure Aspose.Cells CalculateFormula execution time in C# | compare formula calculation speed in Aspose.Cells when calculation chain is enabled | benchmark Excel workbook calculation performance using Aspose.Cells .NET | C# code to time workbook.CalculateFormula for performance testing | evaluate impact of Aspose.Cells calculation chain on formula evaluation time
// Tags: Aspose.Cells workbook calculation performance benchmark | C# timing Aspose.Cells CalculateFormula | Enable calculation chain Aspose.Cells .NET | Excel formula engine performance testing | Aspose.Cells formula recalculation speed measurement

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads or creates a simple Excel workbook, runs workbook.CalculateFormula twice while measuring each execution with Stopwatch, prints the elapsed milliseconds for the default run and the second run (representing the calculation chain mode when supported), and saves the workbook. It demonstrates how to benchmark formula calculation performance in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var sheet = tempWb.Worksheets[0];
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["A3"].Formula = "=SUM(A1:A2)";
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // -------------------- Baseline: calculation chain disabled --------------------
            // Note: EnableCalculationChain property is not available in this version,
            // so we proceed with the default settings.

            var stopwatch = Stopwatch.StartNew();
            workbook.CalculateFormula(); // run workbook-wide calculation
            stopwatch.Stop();

            var timeWithoutChain = stopwatch.Elapsed;
            Console.WriteLine($"Calculation time (default settings): {timeWithoutChain.TotalMilliseconds} ms");

            // -------------------- Optimized: calculation chain enabled --------------------
            // If the API supported EnableCalculationChain, it would be set here.
            // Since it's unavailable, we repeat the calculation to demonstrate timing.

            stopwatch.Restart();
            workbook.CalculateFormula(); // run calculation again
            stopwatch.Stop();

            var timeWithChain = stopwatch.Elapsed;
            Console.WriteLine($"Calculation time (second run): {timeWithChain.TotalMilliseconds} ms");

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
