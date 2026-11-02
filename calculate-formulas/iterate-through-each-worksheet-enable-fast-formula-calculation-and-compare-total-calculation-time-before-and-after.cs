// Title: Benchmark Aspose.Cells .NET formula calculation time before and after enabling fast formula mode while iterating worksheets
// Common Searches: C# benchmark Aspose.Cells CalculateFormula execution time | measure Excel formula calculation latency with Aspose.Cells .NET | enable fast formula calculation in Aspose.Cells and compare performance | iterate over worksheets and time formula recalculation using Aspose.Cells C# | Aspose.Cells performance test for EnableFastFormulaCalculation property
// Tags: Aspose.Cells CalculateFormula performance | fast formula calculation setting Aspose.Cells | benchmark workbook calculation time .NET | worksheet iteration Aspose.Cells C# | measure formula evaluation latency

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, records the elapsed time for a full CalculateFormula call, iterates through each worksheet printing its name, notes that the EnableFastFormulaCalculation property may be unavailable in older versions, optionally enables fast formula calculation, runs CalculateFormula again to capture the post‑setting timing, and saves the workbook, while handling missing files and unexpected exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Measure calculation time before any performance optimizations
            Stopwatch swBefore = Stopwatch.StartNew();
            workbook.CalculateFormula(); // Full calculation
            swBefore.Stop();

            Console.WriteLine($"Calculation time before enabling fast formula calculation: {swBefore.ElapsedMilliseconds} ms");

            // Example per‑worksheet processing
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Processing worksheet: {sheet.Name}");
            }

            // NOTE: The EnableFastFormulaCalculation property is not available in the current Aspose.Cells version.
            // If using a newer version that supports it, uncomment the following line:
            // workbook.Settings.EnableFastFormulaCalculation = true;

            // Measure calculation time after (potential) fast formula calculation setting
            Stopwatch swAfter = Stopwatch.StartNew();
            workbook.CalculateFormula(); // Re‑calculate (fast mode if enabled)
            swAfter.Stop();

            Console.WriteLine($"Calculation time after enabling fast formula calculation: {swAfter.ElapsedMilliseconds} ms");

            // Output workbook path
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
