// Title: Aspose.Cells for .NET – Export Worksheet to HTML with Gridlines and Cell Comments
// Description: Shows how to configure HtmlSaveOptions in Aspose.Cells to generate an HTML file that preserves Excel gridlines and renders cell comments (inline or pop‑up). The sample creates a workbook, adds sample data, inserts a comment, enables ExportGridLines and comment export, and saves the active sheet as HTML.
// Keywords: Aspose.Cells HTML export | ExportGridLines true | IsExportComments | ExportCommentsType PrintInPlace | C# export Excel to HTML with comments | .NET gridlines HTML output | Aspose.Cells HtmlSaveOptions example | Excel to web preview | global | USA
// Common Searches: Aspose.Cells export gridlines and comments to HTML | C# HtmlSaveOptions show cell comments in HTML | How to keep Excel gridlines when saving as HTML with Aspose | Export worksheet with comments using Aspose.Cells .NET | HTML preview of Excel with comments and gridlines
// Developer Intent: Generate an HTML representation of an Excel worksheet that includes both visible gridlines and cell comments.
// Use Cases: Create web‑ready reports that look identical to the Excel view, including gridlines and inline comments. | Provide an interactive worksheet preview on a website where users can see comments without opening Excel. | Export only the active sheet for embedding in documentation or intranet portals while retaining visual fidelity.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet to HTML showing gridlines and cell comments inline. | Explain how to set HtmlSaveOptions.ExportGridLines and ExportCommentsType to PrintInPlace for HTML output. | Provide steps to make gridlines visible in a workbook and ensure they appear in the generated HTML file.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to configure HtmlSaveOptions in Aspose.Cells to generate an HTML file that preserves Excel gridlines and renders cell comments (inline or pop‑up). The sample creates a workbook, adds sample data, inserts a comment, enables ExportGridLines and comment export, and saves the active sheet as HTML.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Make gridlines visible in the worksheet (optional, but mirrors the ExportGridLines setting)
            sheet.IsGridlinesVisible = true;

            // Add some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // Add a comment to a cell
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Configure HTML save options to export both gridlines and comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,          // Enable gridlines in the exported HTML
                IsExportComments = true,         // Export comments (legacy flag)
                // Alternatively, you can specify the comment export type:
                // ExportCommentsType = PrintCommentsType.PrintInPlace,
                ExportActiveWorksheetOnly = true // Export only the active sheet for simplicity
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedWithGridlinesAndComments.html", htmlOptions);

            Console.WriteLine("HTML export completed with gridlines and comments.");
        }
    }
}
