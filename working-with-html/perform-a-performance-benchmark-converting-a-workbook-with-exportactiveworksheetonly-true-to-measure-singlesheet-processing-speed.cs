// Title: Aspose.Cells C# Benchmark: HTML Export of Active Worksheet vs Full Workbook
// Description: A C# program that creates a workbook with two large worksheets, fills them with dummy data, and uses HtmlSaveOptions (ExportActiveWorksheetOnly = true, optional ExportSingleTab) to measure the time required to save only the active sheet to HTML. It then repeats the save with ExportActiveWorksheetOnly disabled to compare the duration for exporting the entire workbook.
// Keywords: Aspose.Cells | C# HTML export benchmark | ExportActiveWorksheetOnly | single sheet HTML conversion | performance measurement Aspose.Cells | HtmlSaveOptions timing | ExportSingleTab | Excel to HTML speed
// Common Searches: Aspose.Cells benchmark HTML export single sheet | measure ExportActiveWorksheetOnly performance C# | HTML export speed Aspose.Cells active worksheet | compare active sheet vs all sheets export time | C# code to time Aspose.Cells HTML save
// Developer Intent: Measure and compare the time required to export only the active worksheet versus the entire workbook to HTML using Aspose.Cells.
// Use Cases: Identify the fastest export configuration for large Excel files | Validate performance gains of ExportActiveWorksheetOnly and ExportSingleTab | Provide timing data for reporting or monitoring tools | Optimize server‑side Excel‑to‑HTML conversion pipelines
// AI Prompts: Write a C# script that logs both execution time and memory consumption when exporting a single worksheet to HTML with Aspose.Cells. | Suggest additional HtmlSaveOptions that can further reduce HTML export time for large worksheets. | Create an automated test that asserts the active‑sheet export completes at least 30% faster than exporting all sheets. | Explain how to integrate the benchmark into a CI pipeline for continuous performance tracking.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // A C# program that creates a workbook with two large worksheets, fills them with dummy data, and uses HtmlSaveOptions (ExportActiveWorksheetOnly = true, optional ExportSingleTab) to measure the time required to save only the active sheet to HTML. It then repeats the save with ExportActiveWorksheetOnly disabled to compare the duration for exporting the entire workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sheet2");

            // Populate both sheets with a large amount of data to make the conversion noticeable
            PopulateSheet(workbook.Worksheets[0], 2000, 50); // Sheet1
            PopulateSheet(workbook.Worksheets[1], 2000, 50); // Sheet2

            // Set the first worksheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                ExportSingleTab = true // optional, improves output for single‑sheet files
            };

            // Measure the time taken to save the active sheet only
            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save("ActiveSheetOnly.html", saveOptions);
            sw.Stop();

            Console.WriteLine($"Time to export active worksheet only: {sw.ElapsedMilliseconds} ms");

            // For comparison, export the whole workbook (all sheets)
            saveOptions.ExportActiveWorksheetOnly = false;
            sw.Restart();
            workbook.Save("AllSheets.html", saveOptions);
            sw.Stop();

            Console.WriteLine($"Time to export all worksheets: {sw.ElapsedMilliseconds} ms");
        }

        // Helper method to fill a worksheet with dummy data
        private static void PopulateSheet(Worksheet sheet, int rows, int columns)
        {
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                }
            }
        }
    }
}
