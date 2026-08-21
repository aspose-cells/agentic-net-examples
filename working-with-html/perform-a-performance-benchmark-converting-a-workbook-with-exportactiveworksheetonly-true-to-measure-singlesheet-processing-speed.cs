// Title: C# Benchmark of Aspose.Cells HTML Export – Active Worksheet Only (ExportActiveWorksheetOnly)
// Description: A console program creates a 5,000‑row by 50‑column workbook, adds a second sheet, sets the first sheet active, and uses HtmlSaveOptions with ExportActiveWorksheetOnly = true (and ExportSingleTab) to time the HTML export of a single sheet. It then repeats the save with ExportActiveWorksheetOnly = false to compare full‑workbook export speed.
// Keywords: Aspose.Cells HTML export benchmark | ExportActiveWorksheetOnly performance | C# Aspose.Cells timing | single sheet HTML save speed | Aspose.Cells .NET HTMLSaveOptions
// Common Searches: Aspose.Cells benchmark active worksheet HTML export | measure Aspose.Cells HTML save time C# | ExportActiveWorksheetOnly vs full workbook speed | how fast is Aspose.Cells HTML export for one sheet | C# performance test Aspose.Cells HtmlSaveOptions
// Developer Intent: Determine the execution time required to export only the active worksheet to HTML with Aspose.Cells and compare it against exporting the entire workbook.
// Use Cases: Assess whether ExportActiveWorksheetOnly reduces export latency for large workbooks. | Validate that the generated HTML contains only the active sheet. | Provide data for selecting optimal HTML export settings in high‑throughput .NET applications.
// AI Prompts: Write a C# loop that runs the single‑sheet HTML export ten times and returns the average elapsed milliseconds using Aspose.Cells. | Suggest ways to accelerate Aspose.Cells HTML export for large worksheets, including alternative options or multi‑threading techniques. | Create a sample report that compares ExportActiveWorksheetOnly true vs false across workbook sizes of 1k, 5k, and 10k rows.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBenchmark
{
    // A console program creates a 5,000‑row by 50‑column workbook, adds a second sheet, sets the first sheet active, and uses HtmlSaveOptions with ExportActiveWorksheetOnly = true (and ExportSingleTab) to time the HTML export of a single sheet. It then repeats the save with ExportActiveWorksheetOnly = false to compare full‑workbook export speed.
    class Program
    {
        static void Main()
        {
            // Create a workbook with a large amount of data to simulate a realistic scenario
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            const int rows = 5000;
            const int cols = 50;

            // Populate the worksheet with sample data
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                }
            }

            // Add a second sheet to ensure ExportActiveWorksheetOnly actually limits the export
            Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");
            secondSheet.Cells["A1"].PutValue("This sheet should NOT be exported.");

            // Set the first sheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                ExportSingleTab = true // optional, improves output when only one sheet is exported
            };

            // Warm‑up run (not measured) to mitigate JIT overhead
            workbook.Save("warmup.html", saveOptions);

            // Measure the time taken to save the workbook with the active‑sheet‑only option
            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save("single_sheet_output.html", saveOptions);
            sw.Stop();

            Console.WriteLine($"Time to export active worksheet only: {sw.ElapsedMilliseconds} ms");

            // For comparison, measure exporting the whole workbook
            HtmlSaveOptions fullSaveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false
            };

            sw.Restart();
            workbook.Save("full_workbook_output.html", fullSaveOptions);
            sw.Stop();

            Console.WriteLine($"Time to export full workbook: {sw.ElapsedMilliseconds} ms");
        }
    }
}
