// Title: Measure multi‑threaded formula calculation performance on a large Excel workbook using Aspose.Cells CalculationOptions in C#
// Common Searches: how to benchmark Aspose.Cells formula calculation speed in C# | enable and configure multi‑threaded calculation in Aspose.Cells .NET | measure execution time of workbook.CalculateFormula for large Excel files | set number of calculation threads using workbook.Settings.NumberOfThreads | performance testing of Aspose.Cells multi‑threaded formula evaluation
// Tags: Aspose.Cells multi-threaded formula calculation | CalculationOptions thread count configuration | Workbook.CalculateFormula performance measurement | large Excel workbook processing with Aspose.Cells | Stopwatch timing of Aspose.Cells calculations

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads a large Excel workbook with Aspose.Cells, relies on the library's default multi‑threaded calculation (adjustable via workbook.Settings.NumberOfThreads), measures the time taken by workbook.CalculateFormula using a Stopwatch, and saves the calculated file. It includes file existence checks and basic error handling.
class MultiThreadedCalculationDemo
{
    static void Main()
    {
        // Paths to the workbook files
        string inputPath = @"C:\Data\LargeWorkbook.xlsx";
        string outputPath = @"C:\Data\LargeWorkbook_Calculated.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Multi‑threaded calculation is enabled by default in recent Aspose.Cells versions.
            // If needed, you can control the number of threads via workbook.Settings.NumberOfThreads.

            // Measure calculation performance
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Perform full workbook calculation
            workbook.CalculateFormula();

            sw.Stop();
            Console.WriteLine($"Multi‑threaded calculation completed in {sw.Elapsed.TotalSeconds:F2} seconds.");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the calculated workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Calculated workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
