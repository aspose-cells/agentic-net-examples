// Title: Aspose.Cells .NET: Export Worksheet to HTML with Gridlines Visible
// Description: This example creates a workbook, enables worksheet gridlines, fills a few cells, and saves the file as HTML using HtmlSaveOptions with ExportGridLines set to true. The resulting HTML preserves the Excel gridline layout for accurate web rendering.
// Keywords: Aspose.Cells | .NET | C# | HTML export | ExportGridLines | Worksheet.IsGridlinesVisible | HtmlSaveOptions | Excel to HTML | gridline rendering | web report
// Common Searches: Aspose.Cells export gridlines to HTML | Enable Excel gridlines in HTML output .NET | HtmlSaveOptions ExportGridLines not working | How to keep gridlines when converting Excel to HTML | C# code for HTML export with visible gridlines
// Developer Intent: Generate an HTML file from a workbook that displays the worksheet’s gridlines.
// Use Cases: Render Excel‑style reports on a website while preserving cell borders. | Provide an online review of data‑entry sheets with the same visual grid as the original workbook. | Create printable HTML invoices or catalogs that retain the familiar gridline layout.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to HTML with gridlines visible and ensure the output folder exists. | Explain why both Worksheet.IsGridlinesVisible and HtmlSaveOptions.ExportGridLines must be set for gridlines to appear in the HTML file. | Show how to change the color or thickness of gridlines when exporting to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, enables worksheet gridlines, fills a few cells, and saves the file as HTML using HtmlSaveOptions with ExportGridLines set to true. The resulting HTML preserves the Excel gridline layout for accurate web rendering.
    public class ExportGridLinesToHtmlDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure gridlines are visible in the worksheet
            worksheet.IsGridlinesVisible = true;

            // Add some sample data so the gridlines can be observed
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(15);

            // Configure HTML save options to export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Enable gridline export
                ExportActiveWorksheetOnly = true      // Export only the active sheet (optional)
            };

            // Define output file path
            string outputPath = "GridLinesExported.html";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML with gridlines visible
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine("HTML file saved with gridlines exported.");
        }
    }
}
