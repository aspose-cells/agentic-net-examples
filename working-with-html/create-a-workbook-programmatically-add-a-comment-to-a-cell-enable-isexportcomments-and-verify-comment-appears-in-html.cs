// Title: Export Cell Comments to HTML with Aspose.Cells for .NET (C#)
// Description: Programmatically creates a workbook, adds a comment to a cell, enables HtmlSaveOptions.IsExportComments, saves the file as HTML, and verifies that the comment text appears in the generated HTML.
// Keywords: Aspose.Cells | C# HTML export | export cell comments | HtmlSaveOptions IsExportComments | save workbook as HTML | verify comment in HTML | Aspose.Cells comment export | programmatic workbook creation | cell comment verification | HTML output with notes
// Common Searches: Aspose.Cells export comments to HTML C# | How to include cell comments when saving as HTML using Aspose.Cells | Enable IsExportComments in HtmlSaveOptions | Check if comment appears in generated HTML Aspose.Cells | C# code to add comment and export to HTML
// Developer Intent: Create an HTML version of an Excel workbook that retains cell comments.
// Use Cases: Generate web‑ready reports with embedded Excel comments for end‑users. | Automate quality checks that confirm comments are present after HTML conversion. | Integrate Excel‑to‑HTML conversion into documentation pipelines while preserving annotations. | Provide interactive tutorials where comment tooltips appear in the HTML view.
// AI Prompts: Write C# code using Aspose.Cells to add a comment to cell B2, enable IsExportComments, and save the workbook as HTML. | Create a reusable method that takes a workbook file path, adds a comment to a given cell, exports to HTML, and returns true if the comment text is found in the output. | Explain the effect of HtmlSaveOptions.IsExportComments on the generated HTML and describe the HTML elements that represent exported comments. | Generate a PowerShell script that calls a .NET assembly to add comments and export to HTML with comments included.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentExportDemo
{
    // Programmatically creates a workbook, adds a comment to a cell, enables HtmlSaveOptions.IsExportComments, saves the file as HTML, and verifies that the comment text appears in the generated HTML.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some data to the cell (optional, just for context)
            worksheet.Cells["A1"].PutValue("Sample Data");

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a test comment";

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true // Enable comment export
            };

            // Save the workbook as HTML with comments exported
            string htmlPath = "output_with_comments.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the comment appears in the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);
            bool commentFound = htmlContent.Contains("This is a test comment");

            Console.WriteLine($"Comment exported to HTML: {commentFound}");
        }
    }
}
