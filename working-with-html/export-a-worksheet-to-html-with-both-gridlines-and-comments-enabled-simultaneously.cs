// Title: Export Excel Worksheet to HTML with Gridlines and Comments using Aspose.Cells for .NET
// Description: Shows how to save a workbook as HTML while keeping visible gridlines and cell comments. The sample sets IsGridlinesVisible, adds data and a comment, configures HtmlSaveOptions (ExportGridLines, IsExportComments), and writes the result to output.html.
// Keywords: Aspose.Cells | C# HtmlSaveOptions | ExportGridLines | IsExportComments | export HTML with gridlines | export HTML with comments | save workbook as HTML | Excel to HTML .NET | gridlines visibility | cell comments export | Aspose.Cells example
// Common Searches: Aspose.Cells export HTML gridlines C# | How to include cell comments when saving Excel as HTML | HtmlSaveOptions ExportGridLines and IsExportComments example | C# convert worksheet to HTML with gridlines and comments | Aspose.Cells HTML output preserving comments
// Developer Intent: Generate an HTML file from a worksheet that displays both gridlines and cell comments.
// Use Cases: Web‑based reports that need the original Excel grid layout and comment notes. | Documentation pages where comments provide extra context beside data cells. | Online spreadsheet preview that retains visual fidelity, including gridlines and comment pop‑ups.
// AI Prompts: Modify the HtmlSaveOptions to also export formulas while keeping gridlines and comments. | Provide a C# snippet that saves each worksheet in a workbook to separate HTML files with gridlines and comments enabled. | Explain how to style exported comments in the HTML output using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to save a workbook as HTML while keeping visible gridlines and cell comments. The sample sets IsGridlinesVisible, adds data and a comment, configures HtmlSaveOptions (ExportGridLines, IsExportComments), and writes the result to output.html.
class ExportGridlinesAndComments
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Make gridlines visible in the worksheet
        worksheet.IsGridlinesVisible = true;

        // Add some sample data
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["B2"].PutValue(123);

        // Add a comment to a cell
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "This is a test comment";

        // Configure HTML save options to export both gridlines and comments
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportGridLines = worksheet.IsGridlinesVisible, // export gridlines
            IsExportComments = true                         // export comments
        };

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", options);
    }
}
