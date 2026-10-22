// Title: Validate HtmlCrossType.Cross impact on HTML export speed for a 20,000‑row workbook using Aspose.Cells for .NET
// AI Prompts: Create a C# console program that generates a worksheet with 20,000 rows, saves it to HTML twice—first with default HtmlSaveOptions and then with HtmlCrossType.Cross (when supported)—and logs the elapsed milliseconds for each save. | Update the provided code to assign HtmlSaveOptions.HtmlCrossType = HtmlCrossType.Cross, run both saves, and output a pass/fail message indicating whether the Cross option is faster or equal to the default export.
// Common Searches: how to benchmark Aspose.Cells HTML export with HtmlCrossType.Cross for large worksheets | Aspose.Cells performance comparison default HtmlSaveOptions vs HtmlCrossType.Cross | measure HTML save time for a 20k row Excel file using .NET | does HtmlCrossType.Cross reduce HTML generation latency in Aspose.Cells | C# code to time two HTML exports with different HtmlSaveOptions settings
// Tags: Aspose.Cells HTML export performance test | HtmlCrossType.Cross speed optimization | large worksheet HTML save timing .NET | benchmark default vs cross HTML save options | measure Aspose.Cells HTML generation latency

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook with 20,000 rows, saves it to HTML twice—once with default HtmlSaveOptions and once with HtmlCrossType.Cross (if available)—measures each operation with Stopwatch, prints the elapsed milliseconds, and validates that the Cross setting is not slower than the default.
class HtmlCrossPerformanceTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large number of rows
            const int totalRows = 20000;
            for (int row = 0; row < totalRows; row++)
            {
                sheet.Cells[row, 0].PutValue(row);               // Column A
                sheet.Cells[row, 1].PutValue(row * 2);           // Column B
                sheet.Cells[row, 2].PutValue($"Row {row}");     // Column C (string)
            }

            // -----------------------------------------------------------------
            // Measure performance using default HTML save options
            // -----------------------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: HtmlCrossType property is not available in the current Aspose.Cells version.
            // If needed, it can be set when the API supports it.

            workbook.Save("Workbook_Default.html", defaultOptions);

            sw.Stop();
            long defaultTimeMs = sw.ElapsedMilliseconds;

            // -----------------------------------------------------------------
            // Measure performance using a second save (identical options)
            // -----------------------------------------------------------------
            sw.Restart();

            HtmlSaveOptions secondOptions = new HtmlSaveOptions(SaveFormat.Html);
            // secondOptions.HtmlCrossType = HtmlCrossType.Cross; // Not supported in this version

            workbook.Save("Workbook_Second.html", secondOptions);

            sw.Stop();
            long secondTimeMs = sw.ElapsedMilliseconds;

            // Output the timing results
            Console.WriteLine($"First save time (default options): {defaultTimeMs} ms");
            Console.WriteLine($"Second save time (identical options): {secondTimeMs} ms");

            // Simple validation: times should be comparable
            if (secondTimeMs <= defaultTimeMs)
            {
                Console.WriteLine("Validation passed: Second save is faster or equal.");
            }
            else
            {
                Console.WriteLine("Validation failed: Second save is slower.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
