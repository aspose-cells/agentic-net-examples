// Title: Aspose.Cells .NET Benchmark: ExportAllSheets Performance for 500 Worksheets to CSV
// Description: Creates a workbook with 500 sheets, sets TxtSaveOptions.ExportAllSheets = true, measures the time required to save the workbook as a single CSV file, and outputs the elapsed milliseconds. Ideal for evaluating scalability and speed of multi‑sheet CSV export in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | ExportAllSheets | CSV export benchmark | 500 worksheets | performance testing | C# | TxtSaveOptions | large workbook scalability | Stopwatch timing | multi‑sheet export speed
// Common Searches: Aspose.Cells export 500 sheets to CSV performance | ExportAllSheets benchmark .NET | how fast is CSV export with many worksheets in Aspose.Cells | measure Aspose.Cells multi‑sheet export time | scalability test for Aspose.Cells CSV output
// Developer Intent: Determine how long Aspose.Cells needs to export a 500‑sheet workbook to a single CSV file when ExportAllSheets is enabled.
// Use Cases: Validate scalability of CSV export for large workbooks. | Compare execution time with ExportAllSheets set to false. | Identify performance bottlenecks before deploying to production. | Provide baseline metrics for optimization of multi‑sheet exports.
// AI Prompts: Analyze the benchmark results and suggest code or configuration changes to reduce export time. | Create a comparable benchmark that exports the same workbook to HTML using HtmlSaveOptions and records both duration and memory usage. | Propose a parallel‑processing approach for sheet preparation while keeping ExportAllSheets behavior intact.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a workbook with 500 sheets, sets TxtSaveOptions.ExportAllSheets = true, measures the time required to save the workbook as a single CSV file, and outputs the elapsed milliseconds. Ideal for evaluating scalability and speed of multi‑sheet CSV export in Aspose.Cells for .NET.
class BenchmarkExportAllSheets
{
    static void Main()
    {
        // Create a new workbook (default format is XLSX)
        Workbook workbook = new Workbook();

        // Add 500 worksheets and put a sample value in each
        for (int i = 1; i <= 500; i++)
        {
            if (i == 1)
            {
                // The first worksheet already exists
                workbook.Worksheets[0].Name = $"Sheet{i}";
                workbook.Worksheets[0].Cells["A1"].PutValue($"Data in Sheet{i}");
            }
            else
            {
                Worksheet ws = workbook.Worksheets.Add($"Sheet{i}");
                ws.Cells["A1"].PutValue($"Data in Sheet{i}");
            }
        }

        // Configure TxtSaveOptions to export all sheets to a CSV file
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.ExportAllSheets = true; // Use the ExportAllSheets property rule

        // Start timing the export operation
        Stopwatch sw = Stopwatch.StartNew();

        // Save the workbook using the configured options
        workbook.Save("BenchmarkAllSheets.csv", saveOptions);

        // Stop timing
        sw.Stop();

        Console.WriteLine($"Exported 500 worksheets to CSV in {sw.ElapsedMilliseconds} ms.");

        // Release resources
        workbook.Dispose();
    }
}
