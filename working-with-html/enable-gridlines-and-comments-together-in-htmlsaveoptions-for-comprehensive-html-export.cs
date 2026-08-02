// Title: Export Excel to HTML with Gridlines and Comments using Aspose.Cells for .NET
// Description: Demonstrates how to save an Aspose.Cells workbook as HTML while preserving worksheet gridlines and cell comments. The example sets ExportGridLines and IsExportComments in HtmlSaveOptions, creates sample data, adds a comment, and generates an HTML file that displays both borders and comment tooltips.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | ExportGridLines | IsExportComments | HTML export Excel | gridlines in HTML | cell comments HTML | Excel to web view
// Common Searches: Aspose.Cells export gridlines to HTML | How to include cell comments when saving Excel as HTML | HtmlSaveOptions ExportGridLines and IsExportComments example | C# export workbook to HTML with comments | Enable gridlines and comments in Aspose.Cells HTML output
// Developer Intent: Generate an HTML representation of an Excel workbook that shows both the worksheet gridlines and any cell comments.
// Use Cases: Web‑based reporting portals that need the original Excel layout with annotation pop‑ups. | Documentation sites where reviewers must see cell borders and author notes without opening Excel. | Internal dashboards that display spreadsheet data with visual grid structure and comment tooltips for context.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook to HTML with gridlines and comments enabled, and customize the comment style with CSS. | Explain the interaction between ExportGridLines and IsExportComments in HtmlSaveOptions and how to toggle them per worksheet. | Show how to export multiple worksheets into a single HTML file while keeping gridlines and comments visible for each sheet.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to save an Aspose.Cells workbook as HTML while preserving worksheet gridlines and cell comments. The example sets ExportGridLines and IsExportComments in HtmlSaveOptions, creates sample data, adds a comment, and generates an HTML file that displays both borders and comment tooltips.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Make gridlines visible in the worksheet (optional, but reflects the source setting)
            sheet.IsGridlinesVisible = true;

            // Add some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Add a comment to a cell
            int commentIndex = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Fresh and crispy";

            // Configure HTML save options to export both gridlines and comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,      // Enable gridlines in the exported HTML
                IsExportComments = true      // Enable comments in the exported HTML
            };

            // Save the workbook as HTML with the specified options
            string outputPath = "ExportWithGridlinesAndComments.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
