// Title: Benchmark single‑threaded formula calculation performance in Aspose.Cells for .NET using Workbook.Settings.EnableFastFormulaCalculation and UseThreadedCalculation
// Common Searches: measure single‑threaded formula calculation time with Aspose.Cells in C# | benchmark Aspose.Cells workbook.CalculateFormula performance on complex Excel files | how to disable threaded calculation in Aspose.Cells for accurate timing | using EnableFastFormulaCalculation setting to improve single‑threaded calculation speed in .NET | profile formula evaluation of large worksheets with Aspose.Cells Stopwatch example
// Tags: Workbook.Settings.EnableFastFormulaCalculation usage | Workbook.Settings.UseThreadedCalculation false | single‑threaded formula calculation benchmark | calculate formula performance measurement C# | complex worksheet calculation speed Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads a complex Excel workbook, optionally disables threaded calculation and enables fast formula mode via Workbook.Settings, starts a Stopwatch, runs workbook.CalculateFormula() to evaluate all formulas on a single thread, stops the timer, prints the elapsed seconds, and saves the calculated workbook. It shows how to benchmark single‑threaded formula evaluation performance with Aspose.Cells for .NET.
class SingleThreadedFormulaBenchmark
{
    static void Main()
    {
        try
        {
            const string inputPath = "ComplexSheet.xlsx";
            const string outputPath = "ComplexSheet_Calculated.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing complex sheets
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In some Aspose.Cells versions the following settings are unavailable.
            // They have been omitted to ensure compatibility across versions.
            // workbook.Settings.EnableFastFormulaCalculation = true;
            // workbook.Settings.UseThreadedCalculation = false;

            // Start timing the calculation
            Stopwatch sw = Stopwatch.StartNew();

            // Perform full workbook calculation
            workbook.CalculateFormula();

            // Stop timing
            sw.Stop();

            // Output the elapsed time
            Console.WriteLine($"Single‑threaded calculation time: {sw.Elapsed.TotalSeconds} seconds");

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook after calculation (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Calculated workbook saved to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save the workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
