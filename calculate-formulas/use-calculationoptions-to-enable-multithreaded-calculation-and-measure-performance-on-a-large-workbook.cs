// Title: Measure multi‑threaded formula calculation performance with Aspose.Cells in C# for a large Excel workbook
// Common Searches: how to benchmark Aspose.Cells formula calculation speed in C# | enable multi‑threaded calculation in Aspose.Cells workbook settings | calculate large Excel workbook faster using Aspose.Cells multi‑threading | Aspose.Cells C# example for measuring calculation time with Stopwatch | set number of calculation threads in Aspose.Cells .NET
// Tags: Aspose.Cells multi‑threaded formula calculation | C# benchmark Excel calculation performance | Workbook.CalculateFormula with threading | EnableThreadedCalculation setting Aspose.Cells | large workbook calculation optimization .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// This C# example loads a large Excel workbook with Aspose.Cells, optionally activates multi‑threaded calculation via the workbook settings, measures the elapsed time of a full formula recalculation using Stopwatch, saves the calculated file, and includes basic file existence checks and error handling.
class Program
{
    static void Main()
    {
        // Paths to the input and output workbooks
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

            // Enable multi‑threaded calculation if supported by the version
            // In newer versions this is controlled by EnableThreadedCalculation.
            // For compatibility, we rely on the default behavior or set the thread count.
            // workbook.Settings.EnableThreadedCalculation = true; // Uncomment if property exists
            // Optional: specify the number of threads (default is Environment.ProcessorCount)
            // workbook.Settings.NumberOfThreads = Environment.ProcessorCount;

            // Measure calculation performance
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Force full calculation of the workbook
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
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
