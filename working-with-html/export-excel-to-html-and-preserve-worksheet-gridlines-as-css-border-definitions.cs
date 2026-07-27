// Title: Export Excel to HTML with Gridlines as CSS Borders using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, enable gridlines, add sample data, and use HtmlSaveOptions (ExportGridLines = true, ExportActiveWorksheetOnly = true) to save the worksheet as an HTML file where Excel gridlines are rendered as CSS border definitions.
// Keywords: Aspose.Cells HTML export | ExportGridLines property | Excel gridlines to HTML | CSS borders from Excel | .NET workbook to HTML | Preserve Excel layout in web
// Common Searches: Aspose.Cells keep gridlines when exporting to HTML | HtmlSaveOptions ExportGridLines example C# | Export active worksheet only with gridlines Aspose | Render Excel gridlines as CSS borders
// Developer Intent: Generate an HTML representation of an Excel worksheet that retains the original gridline appearance using CSS borders.
// Use Cases: Web‑based reporting that mirrors Excel’s gridline layout | Embedding a single worksheet in a web page while preserving visual fidelity | Creating printable HTML snapshots of Excel data with clear cell separation
// AI Prompts: Show how to customize the CSS style of exported gridlines with HtmlSaveOptions in Aspose.Cells. | Provide an example that exports multiple worksheets to separate HTML files, each preserving gridlines. | Explain how to turn off gridline export and apply custom cell borders when saving to HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsExportGridlines
{
    // Demonstrates how to create a workbook, enable gridlines, add sample data, and use HtmlSaveOptions (ExportGridLines = true, ExportActiveWorksheetOnly = true) to save the worksheet as an HTML file where Excel gridlines are rendered as CSS border definitions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Make gridlines visible in the worksheet (optional, but mirrors Excel UI)
            sheet.IsGridlinesVisible = true;

            // Add some sample data to demonstrate the gridlines
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // Configure HTML save options to export gridlines as CSS borders
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Preserve gridlines
                ExportActiveWorksheetOnly = true      // Export only the active sheet (optional)
            };

            // Save the workbook as an HTML file with gridlines preserved
            string outputPath = "WorkbookWithGridlines.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
