// Title: Measure and Log Aspose.Cells Workbook Processing Times with C# Stopwatch
// Description: C# example that loads an Excel workbook using Aspose.Cells, then uses System.Diagnostics.Stopwatch to record the duration of loading, formula calculation, custom cell updates, and saving. The logged timings help pinpoint performance bottlenecks in Excel processing pipelines.
// Keywords: Aspose.Cells performance logging | C# Stopwatch Excel timing | measure workbook load time | formula calculation duration | save operation latency | Excel processing profiling .NET | identify performance bottlenecks | Aspose.Cells benchmark | Excel automation timing
// Common Searches: how to profile Aspose.Cells workbook operations in C# | log load and save time for Excel files using Aspose.Cells | measure formula calculation speed with Aspose.Cells | C# stopwatch example for Excel processing | performance monitoring Aspose.Cells workbook
// Developer Intent: Track the execution time of each major Aspose.Cells workbook step to discover and address slow‑running operations.
// Use Cases: Detect slow file parsing when loading large workbooks. | Identify heavy or complex formulas that increase calculation time. | Benchmark custom data‑transformation logic before saving. | Compare I/O performance across different storage locations.
// AI Prompts: Create a reusable C# helper that logs Aspose.Cells step durations to a CSV file with timestamps. | Show how to integrate asynchronous performance logging into an existing Aspose.Cells processing pipeline. | Generate code that automatically highlights the longest step (load, calculate, custom processing, save) and suggests optimization tips.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceLogging
{
    // C# example that loads an Excel workbook using Aspose.Cells, then uses System.Diagnostics.Stopwatch to record the duration of loading, formula calculation, custom cell updates, and saving. The logged timings help pinpoint performance bottlenecks in Excel processing pipelines.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: process a workbook and log each step's duration.
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "ProcessedWorkbook.xlsx";

            ProcessWorkbook(inputPath, outputPath);
        }

        /// <param name="inputFile">Path to the source workbook.</param>
        /// <param name="outputFile">Path where the processed workbook will be saved.</param>
        static void ProcessWorkbook(string inputFile, string outputFile)
        {
            // Stopwatch for measuring durations.
            Stopwatch sw = new Stopwatch();

            // ------------------- Create / Load -------------------
            sw.Start();
            // LoadOptions can be used if you need to set interrupt monitors, etc.
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputFile, loadOptions);
            sw.Stop();
            Console.WriteLine($"Load time: {sw.ElapsedMilliseconds} ms");
            sw.Reset();

            // ------------------- Optional Processing -------------------
            // Example: calculate all formulas.
            sw.Start();
            workbook.CalculateFormula();
            sw.Stop();
            Console.WriteLine($"Formula calculation time: {sw.ElapsedMilliseconds} ms");
            sw.Reset();

            // Add any additional processing here and log its duration similarly.
            // For demonstration, we'll just add a simple value.
            sw.Start();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Processed at " + DateTime.Now);
            sw.Stop();
            Console.WriteLine($"Custom processing time: {sw.ElapsedMilliseconds} ms");
            sw.Reset();

            // ------------------- Save -------------------
            sw.Start();
            workbook.Save(outputFile);
            sw.Stop();
            Console.WriteLine($"Save time: {sw.ElapsedMilliseconds} ms");
        }
    }
}
