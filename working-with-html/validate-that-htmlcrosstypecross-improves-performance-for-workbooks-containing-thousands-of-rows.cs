// Title: HtmlCrossType.Cross vs Default: Speed Test for Exporting a 100k‑row Workbook to HTML with Aspose.Cells .NET
// Description: Generates a workbook of 100,000 rows and 5 columns filled with long strings, saves it twice—once using HtmlCrossStringType.Default and once using HtmlCrossStringType.Cross—while measuring the elapsed time for each export.
// Keywords: Aspose.Cells | HtmlCrossType | Cross mode | Default mode | HTML export performance | large workbook | C# benchmark | Excel to HTML conversion | rendering speed | .NET
// Common Searches: Aspose.Cells HtmlCrossType performance test | HTML export speed for 100k rows Aspose.Cells | Does HtmlCrossStringType.Cross speed up large Excel to HTML conversion | Benchmark HtmlCrossStringType options in .NET | Fast HTML rendering of massive worksheets with Aspose.Cells
// Developer Intent: Find out if switching to the Cross setting shortens the time required to produce HTML from a very large worksheet compared with the default configuration.
// Use Cases: Select the best cross‑string handling mode for high‑volume reporting dashboards. | Meet latency targets in web apps that serve extensive spreadsheets as HTML. | Compare memory and CPU usage of different HtmlCrossStringType values during bulk export.
// AI Prompts: Write a C# program that creates a workbook with 150,000 rows, exports it to HTML using both HtmlCrossStringType.Default and HtmlCrossStringType.Cross, and prints a side‑by‑side performance summary. | Explain the internal differences between the Default and Cross modes for handling overflow text in Aspose.Cells HTML conversion. | Recommend additional HtmlSaveOptions (e.g., ExportActiveWorksheetOnly, MaxRowsPerPage) that can further accelerate HTML export of very large workbooks.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlCrossTypePerformanceDemo
{
    // Generates a workbook of 100,000 rows and 5 columns filled with long strings, saves it twice—once using HtmlCrossStringType.Default and once using HtmlCrossStringType.Cross—while measuring the elapsed time for each export.
    class Program
    {
        static void Main()
        {
            // Number of rows to simulate a large workbook
            const int totalRows = 100_000;
            const int totalColumns = 5;

            // ------------------------------------------------------------
            // Create a workbook and populate it with sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    // Put a string that will normally overflow the cell width
                    cells[row, col].PutValue($"Row{row}_Col{col}_LongTextThatWillCrossCells");
                }
            }

            // ------------------------------------------------------------
            // Save with the default HtmlCrossStringType (Default) and measure time
            // ------------------------------------------------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions
            {
                // Default value is HtmlCrossType.Default; set explicitly for clarity
                HtmlCrossStringType = HtmlCrossType.Default
            };

            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save("LargeWorkbook_Default.html", defaultOptions);
            sw.Stop();
            Console.WriteLine($"Saving with HtmlCrossType.Default took: {sw.ElapsedMilliseconds} ms");

            // ------------------------------------------------------------
            // Save with HtmlCrossStringType.Cross and measure time
            // ------------------------------------------------------------
            HtmlSaveOptions crossOptions = new HtmlSaveOptions
            {
                HtmlCrossStringType = HtmlCrossType.Cross
            };

            sw.Restart();
            workbook.Save("LargeWorkbook_Cross.html", crossOptions);
            sw.Stop();
            Console.WriteLine($"Saving with HtmlCrossType.Cross took: {sw.ElapsedMilliseconds} ms");

            // ------------------------------------------------------------
            // Cleanup
            // ------------------------------------------------------------
            workbook.Dispose();
        }
    }
}
