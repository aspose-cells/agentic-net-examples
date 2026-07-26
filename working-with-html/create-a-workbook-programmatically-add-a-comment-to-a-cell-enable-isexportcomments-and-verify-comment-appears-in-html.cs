// Title: Export Cell Comments to HTML with Aspose.Cells for .NET
// Description: Demonstrates how to add a comment to a worksheet cell, enable HtmlSaveOptions.IsExportComments, save the workbook as HTML, and programmatically verify that the comment appears in the generated file.
// Keywords: Aspose.Cells export comments HTML | HtmlSaveOptions IsExportComments .NET | C# add cell comment Aspose | verify comment in HTML Aspose.Cells | save workbook as HTML with comments
// Common Searches: Aspose.Cells export cell comments to HTML | HtmlSaveOptions IsExportComments example C# | check comment in saved HTML Aspose.Cells | add comment and save workbook as HTML .NET
// Developer Intent: Create an HTML representation of a workbook that includes cell comments and confirm the comment text is present.
// Use Cases: Generate web‑ready reports where explanatory comments are retained as tooltips. | Automate documentation pipelines that convert Excel workbooks to HTML with embedded notes. | Implement unit tests that validate comment export functionality in Aspose.Cells.
// AI Prompts: Write C# code using Aspose.Cells to add a comment to cell B2, export the workbook to HTML with comments, and return true if the comment text is found. | Explain the impact of HtmlSaveOptions.IsExportComments on the HTML output and show how to programmatically verify the comment content. | Create an MSTest method that creates a workbook, adds a comment, saves it as HTML with comments enabled, and asserts the comment string exists in the file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentExportDemo
{
    // Demonstrates how to add a comment to a worksheet cell, enable HtmlSaveOptions.IsExportComments, save the workbook as HTML, and programmatically verify that the comment appears in the generated file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some data to the cell (optional, just for context)
            worksheet.Cells["A1"].PutValue("Cell with comment");

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a test comment";

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true
            };

            // Define output HTML file path
            string htmlPath = "output_with_comments.html";

            // Save the workbook as HTML with comments exported
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the comment appears in the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);
            bool commentFound = htmlContent.Contains("This is a test comment");

            Console.WriteLine($"Comment exported to HTML: {commentFound}");
        }
    }
}
