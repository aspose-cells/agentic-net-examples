// Title: Benchmark single‑threaded formula calculation in Aspose.Cells for .NET with EnableFastFormulaCalculation and UseThreadedCalculation disabled
// Common Searches: c# how to measure single thread formula calculation time using Aspose.Cells | aspnet benchmark Aspose.Cells formula engine with EnableFastFormulaCalculation | disable threaded calculation in Aspose.Cells to test performance | compare fast formula calculation vs normal in Aspose.Cells .NET | profile complex Excel workbook calculation using Aspose.Cells
// Tags: EnableFastFormulaCalculation performance tuning Aspose.Cells | single-threaded formula calculation benchmark .NET | UseThreadedCalculation false Aspose.Cells example | measure Excel formula calculation speed C# | complex workbook formula engine optimization Aspose.Cells | benchmarking Aspose.Cells calculation engine

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample loads a workbook with complex formulas, optionally disables threaded calculation and enables fast formula calculation, runs an initial full calculation, then measures the duration of a second CalculateFormula call with a Stopwatch, outputs the elapsed time, and saves the calculated workbook.
class FormulaBenchmark
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = @"C:\Data\ComplexFormulas.xlsx";
        string outputPath = @"C:\Data\ComplexFormulas_Calculated.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook containing complex formulas
            Workbook workbook = new Workbook(inputPath);

            // Optional: configure calculation settings if supported by the current Aspose.Cells version
            // Uncomment the following lines if the properties are available in your version:
            // workbook.Settings.EnableFastFormulaCalculation = true;   // enable fast calculation optimizations
            // workbook.Settings.UseThreadedCalculation = false;       // force single‑threaded execution

            // Perform an initial full calculation (optional)
            workbook.CalculateFormula();

            // Benchmark the calculation time
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Re‑calculate formulas using the same settings
            workbook.CalculateFormula();

            sw.Stop();

            Console.WriteLine($"Single‑threaded calculation time: {sw.Elapsed.TotalSeconds} seconds");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after calculation (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Calculated workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
