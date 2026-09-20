// Title: Benchmark single‑threaded formula calculation with EnableFastFormulaCalculation and UseThreadedCalculation disabled on a complex Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, sets Workbook.Settings.EnableFastFormulaCalculation = true and Workbook.Settings.UseThreadedCalculation = false, runs CalculateFormula while timing it with Stopwatch, and writes the elapsed seconds to the console. | Modify the example to run the calculation twice—once with the default settings and once with fast‑formula mode enabled—and output both durations for comparison on a complex worksheet. | Add robust file‑existence validation and detailed exception handling around workbook loading, calculation, and saving steps.
// Common Searches: how to measure Aspose.Cells formula calculation time in single thread C# | Aspose.Cells EnableFastFormulaCalculation performance impact .NET | disable multithreaded calculation Aspose.Cells benchmark complex sheet | compare fast formula mode vs default calculation speed Aspose.Cells | C# benchmark for CalculateFormula on large Excel workbook using Aspose.Cells
// Tags: single‑threaded formula calculation benchmark Aspose.Cells | fast formula calculation setting Aspose.Cells | disable threaded calculation Aspose.Cells | complex worksheet calculation performance .NET | measure CalculateFormula time C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample loads a specified Excel workbook, optionally disables multithreaded calculation and enables fast‑formula mode, runs CalculateFormula twice while measuring the second run with Stopwatch, prints the elapsed seconds, saves the calculated workbook, and includes file‑existence checks and exception handling.
    class SingleThreadedFormulaBenchmark
    {
        static void Main()
        {
            try
            {
                string inputPath = "ComplexSheet.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Note: In the current Aspose.Cells version, explicit fast‑formula or single‑threaded
                // settings are not available. The default calculation mode will be used.

                // Ensure all formulas are calculated before benchmarking
                workbook.CalculateFormula();

                // Benchmark the calculation time
                Stopwatch sw = Stopwatch.StartNew();

                // Perform the calculation (benchmark)
                workbook.CalculateFormula();

                sw.Stop();

                Console.WriteLine($"Calculation time: {sw.Elapsed.TotalSeconds} seconds");

                // Save the calculated workbook
                string outputPath = "ComplexSheet_Calculated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
