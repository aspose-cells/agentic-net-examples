// Title: Aspose.Cells .NET Benchmark: Export 500 Worksheets to CSV with ExportAllSheets = true
// Description: C# example that builds a workbook with 500 sheets, enables TxtSaveOptions.ExportAllSheets, times the CSV save operation, and prints the elapsed milliseconds—showing the scalability of Aspose.Cells CSV export.
// Keywords: Aspose.Cells | .NET | C# | CSV export | ExportAllSheets | benchmark | performance test | 500 worksheets | large workbook | save time measurement | TxtSaveOptions
// Common Searches: Aspose.Cells export 500 sheets to CSV performance | CSV export speed Aspose.Cells .NET | ExportAllSheets benchmark | measure Aspose.Cells save time | how fast can Aspose.Cells save large workbook
// Developer Intent: Determine how long Aspose.Cells takes to save a 500‑sheet workbook to CSV when ExportAllSheets is enabled.
// Use Cases: Assess scalability of CSV export for multi‑sheet reports | Validate batch processing time for data‑export pipelines | Compare ExportAllSheets true versus false in large workbooks | Set a performance baseline for CI/CD testing of Aspose.Cells conversions
// AI Prompts: Write a C# script that records CPU and memory while exporting 500 worksheets to CSV with ExportAllSheets true using Aspose.Cells. | Create a loop that runs the conversion 10 times and returns the average duration and standard deviation. | Suggest ways to accelerate CSV export for massive workbooks in Aspose.Cells, such as disabling calculations, using streaming, or parallelizing sheet processing. | Provide PowerShell commands to invoke the compiled benchmark and capture timing metrics.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // C# example that builds a workbook with 500 sheets, enables TxtSaveOptions.ExportAllSheets, times the CSV save operation, and prints the elapsed milliseconds—showing the scalability of Aspose.Cells CSV export.
    class Program
    {
        static void Main()
        {
            // Initialize a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add 500 worksheets and put a simple value in each
            for (int i = 0; i < 500; i++)
            {
                // Add a new worksheet; the first worksheet already exists at index 0
                if (i > 0)
                {
                    workbook.Worksheets.Add($"Sheet{i + 1}");
                }

                // Access the current worksheet and set a sample value
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Cells["A1"].PutValue($"Data in sheet {i + 1}");
            }

            // Prepare TxtSaveOptions with ExportAllSheets = true (feature rule)
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                ExportAllSheets = true
            };

            // Benchmark the conversion (saving) time
            Stopwatch sw = Stopwatch.StartNew();

            // Save the workbook to CSV using the options (lifecycle rule: save)
            workbook.Save("BenchmarkOutput.csv", saveOptions);

            sw.Stop();

            Console.WriteLine($"Converted workbook with 500 worksheets to CSV in {sw.ElapsedMilliseconds} ms.");
            
            // Clean up
            workbook.Dispose();
        }
    }
}
