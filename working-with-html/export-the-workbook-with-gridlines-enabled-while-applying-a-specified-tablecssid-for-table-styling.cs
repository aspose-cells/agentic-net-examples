// Title: Export Aspose.Cells Workbook to HTML with Gridlines and Custom TableCssId (C#)
// Description: Shows how to enable worksheet gridlines, set HtmlSaveOptions.ExportGridLines, assign a TableCssId, and save the workbook as HTML using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | HTML export | ExportGridLines | TableCssId | C# | .NET | workbook to HTML | gridlines visible | custom CSS ID | HTML table styling
// Common Searches: Aspose.Cells export HTML gridlines | HtmlSaveOptions TableCssId example | C# export Excel to HTML with gridlines | keep gridlines when saving as HTML Aspose | apply custom CSS to Aspose.Cells HTML output
// Developer Intent: Generate an HTML file from a workbook that retains Excel gridlines and uses a specified CSS ID for the table.
// Use Cases: Web dashboards that require Excel‑style gridlines for data clarity | Automated report generation that applies corporate branding via a CSS ID | Embedding Excel data in web pages while preserving the original layout | Creating printable HTML reports with consistent table styling
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with ExportGridLines = true and TableCssId = "custom-table-style". | Provide CSS definitions that match the TableCssId used in the exported HTML from Aspose.Cells. | Explain the impact of ExportGridLines and TableCssId on the generated HTML and how to ensure the custom CSS is applied correctly.

using System;
using Aspose.Cells;

namespace AsposeCellsExportWithGridlinesAndTableCss
{
    // Shows how to enable worksheet gridlines, set HtmlSaveOptions.ExportGridLines, assign a TableCssId, and save the workbook as HTML using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable gridlines visibility in the worksheet
            worksheet.IsGridlinesVisible = true;

            // Add some sample data to visualize the gridlines
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["A2"].PutValue("Data 1");
            worksheet.Cells["B2"].PutValue("Data 2");

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export the gridlines as they are visible in the worksheet
                ExportGridLines = true,

                // Apply a custom TableCssId for styling the generated HTML table
                TableCssId = "custom-table-style"
            };

            // Save the workbook as HTML using the configured options (lifecycle save)
            workbook.Save("ExportedWithGridlines.html", saveOptions);
        }
    }
}
