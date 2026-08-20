// Title: Export Worksheet to HTML with Gridlines and Comments – Aspose.Cells for .NET
// Description: Demonstrates how to save a single worksheet as HTML while preserving visible gridlines and cell comments using Aspose.Cells HtmlSaveOptions (ExportGridLines, IsExportComments, ExportActiveWorksheetOnly).
// Keywords: Aspose.Cells HTML export | ExportGridLines | IsExportComments | ExportActiveWorksheetOnly | C# export worksheet to HTML | gridlines in HTML output | cell comments HTML Aspose | save single worksheet as HTML
// Common Searches: Aspose.Cells export worksheet to HTML with gridlines | How to include cell comments in HTML export using Aspose.Cells | C# HtmlSaveOptions gridlines comments | Export only active sheet to HTML Aspose.Cells | HTML output with Excel gridlines and comments
// Developer Intent: Generate an HTML file for the active worksheet that shows both gridlines and any cell comments.
// Use Cases: Create a web‑ready view of a spreadsheet that keeps the original grid layout and comment tooltips. | Produce an HTML report for documentation or email that includes annotation comments. | Embed a single worksheet in a web application while preserving visual fidelity of gridlines and comments.
// AI Prompts: Write C# code with Aspose.Cells to export the active worksheet to HTML, enabling gridlines and comments. | Explain the impact of ExportGridLines, IsExportComments, and ExportActiveWorksheetOnly on the HTML result in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to save a single worksheet as HTML while preserving visible gridlines and cell comments using Aspose.Cells HtmlSaveOptions (ExportGridLines, IsExportComments, ExportActiveWorksheetOnly).
class ExportWorksheetHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure gridlines are visible in the worksheet
        sheet.IsGridlinesVisible = true;

        // Add some sample data
        sheet.Cells["A1"].PutValue("Hello World");
        sheet.Cells["B2"].PutValue(12345);

        // Add a comment to a cell
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a sample comment";

        // Set HTML save options to export both gridlines and comments
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportGridLines = true,          // Export gridlines
            IsExportComments = true,         // Export comments
            ExportActiveWorksheetOnly = true // Export only the active worksheet
        };

        // Save the workbook as HTML
        workbook.Save("output.html", options);
    }
}
