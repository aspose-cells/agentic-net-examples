// Title: Enable Gridlines When Exporting Excel to HTML with Aspose.Cells for .NET
// Description: Learn how to export an Excel workbook to HTML while preserving the worksheet gridlines using Aspose.Cells. The example creates a workbook, adds sample data, sets IsGridlinesVisible, configures HtmlSaveOptions with ExportGridLines = true (and optionally ExportActiveWorksheetOnly), and saves the result as GridlinesEnabled.html.
// Keywords: Aspose.Cells HTML export gridlines | ExportGridLines true C# | Aspose.Cells enable gridlines HTML | HtmlSaveOptions ExportGridLines | C# Excel to HTML with gridlines
// Common Searches: Aspose.Cells enable gridlines in HTML output | HtmlSaveOptions ExportGridLines example | C# export Excel worksheet to HTML with gridlines | How to show Excel gridlines in HTML using Aspose.Cells
// Developer Intent: Generate HTML from an Excel workbook that displays the original gridlines.
// Use Cases: Create a web‑preview of a spreadsheet that retains Excel‑style gridlines for clarity. | Produce printable HTML reports where cell borders must be visible. | Export only the active sheet of a multi‑sheet workbook while keeping gridlines for embedding in web pages.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML with gridlines enabled and exports only the active worksheet. | Explain the interaction between Worksheet.IsGridlinesVisible and HtmlSaveOptions.ExportGridLines when converting Excel to HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsGridlinesHtml
{
    // Learn how to export an Excel workbook to HTML while preserving the worksheet gridlines using Aspose.Cells. The example creates a workbook, adds sample data, sets IsGridlinesVisible, configures HtmlSaveOptions with ExportGridLines = true (and optionally ExportActiveWorksheetOnly), and saves the result as GridlinesEnabled.html.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data so the gridlines are visible in the output
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // Ensure gridlines are visible in the worksheet
            sheet.IsGridlinesVisible = true;

            // Create HTML save options and enable exporting of gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Enable gridlines in the HTML output
                ExportActiveWorksheetOnly = true      // Export only the active worksheet (optional)
            };

            // Save the workbook as HTML with gridlines enabled
            workbook.Save("GridlinesEnabled.html", htmlOptions);

            Console.WriteLine("HTML file saved with gridlines enabled.");
        }
    }
}
