// Title: Measure Excel‑to‑PDF conversion time with Aspose.Cells and C# Stopwatch
// Description: A concise C# example that loads an Excel workbook with Aspose.Cells, uses System.Diagnostics.Stopwatch to time the Save operation to PDF, and writes the elapsed milliseconds and formatted TimeSpan to the console for performance monitoring.
// Keywords: Aspose.Cells conversion timing | C# Stopwatch performance | Excel to PDF benchmark | measure workbook.Save latency | .NET export speed analysis | global performance testing | US developers performance tools
// Common Searches: how to time Aspose.Cells PDF export in C# | measure Excel conversion duration with Stopwatch | benchmark Aspose.Cells save performance | log conversion time for Excel to PDF | performance testing Aspose.Cells workbook.Save
// Developer Intent: The developer needs to capture and log the duration of an Excel‑to‑PDF conversion performed by Aspose.Cells.
// Use Cases: Compare conversion speeds of different SaveFormat options (PDF, XPS, HTML). | Record batch conversion times in logs or monitoring dashboards. | Detect performance regressions after upgrading Aspose.Cells by tracking elapsed time.
// AI Prompts: Create a reusable C# method that accepts a source file path and target format, measures the conversion with Stopwatch, and returns the elapsed TimeSpan. | Show how to write conversion timings to a CSV or JSON log instead of the console. | Explain how to run parallel conversions of multiple workbooks while preserving individual Stopwatch measurements for each task.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // A concise C# example that loads an Excel workbook with Aspose.Cells, uses System.Diagnostics.Stopwatch to time the Save operation to PDF, and writes the elapsed milliseconds and formatted TimeSpan to the console for performance monitoring.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourcePath = "SourceWorkbook.xlsx";

            // Path to the converted output file (e.g., PDF)
            string outputPath = "ConvertedWorkbook.pdf";

            // Load the workbook (this is the operation we want to measure)
            Workbook workbook = new Workbook(sourcePath);

            // Create a Stopwatch instance to measure the conversion duration
            Stopwatch stopwatch = new Stopwatch();

            // Start timing before the conversion begins
            stopwatch.Start();

            // Perform the conversion (saving to PDF in this example)
            workbook.Save(outputPath, SaveFormat.Pdf);

            // Stop timing after the conversion completes
            stopwatch.Stop();

            // Log the elapsed time in milliseconds
            Console.WriteLine($"Conversion completed in {stopwatch.ElapsedMilliseconds} ms.");

            // Optionally, also display the elapsed time in a more readable format
            Console.WriteLine($"Elapsed time: {stopwatch.Elapsed}");
        }
    }
}
