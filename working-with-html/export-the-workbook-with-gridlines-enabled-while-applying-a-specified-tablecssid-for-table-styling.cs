// Title: Export Excel to HTML with Gridlines and Custom TableCssId using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, enables worksheet gridlines, adds sample data, and saves the file as HTML with ExportGridLines=true, a user‑defined TableCssId, and optional ExportActiveWorksheetOnly=true.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | ExportGridLines | TableCssId | ExportActiveWorksheetOnly | Excel to HTML | gridlines in HTML | custom table CSS
// Common Searches: Aspose.Cells export gridlines to HTML | How to set TableCssId in HtmlSaveOptions | Export only active worksheet as HTML Aspose | Enable worksheet gridlines in HTML output C# | Custom CSS ID for Aspose.Cells HTML table
// Developer Intent: Generate an HTML representation of a workbook that shows Excel gridlines and uses a specified TableCssId for styling.
// Use Cases: Render Excel reports on a web page with the same gridline layout as the source file. | Apply site‑wide CSS rules to the exported table by targeting a custom TableCssId. | Export a single worksheet for embedding in a portal while preserving visual fidelity.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to HTML with visible gridlines and a custom TableCssId, and explain each HtmlSaveOptions property. | Show how to link an external stylesheet in the exported HTML while keeping the TableCssId applied to the table. | Provide a step‑by‑step tutorial for exporting multiple worksheets to separate HTML files, each with gridlines and a unique TableCssId.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Creates a workbook, enables worksheet gridlines, adds sample data, and saves the file as HTML with ExportGridLines=true, a user‑defined TableCssId, and optional ExportActiveWorksheetOnly=true.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable gridlines visibility in the worksheet
            worksheet.IsGridlinesVisible = true;

            // Add some sample data to demonstrate the output
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Alice");
            worksheet.Cells["B3"].PutValue(25);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export the gridlines as they are visible in the worksheet
                ExportGridLines = true,

                // Apply a custom TableCssId for styling the generated HTML table
                TableCssId = "custom-table-style",

                // Optional: export only the active worksheet
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save("ExportedWithGridlines.html", htmlOptions);
        }
    }
}
