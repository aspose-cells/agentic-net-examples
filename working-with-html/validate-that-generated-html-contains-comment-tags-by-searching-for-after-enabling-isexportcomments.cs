// Title: Validate HTML Comment Tags Exported with AspNet Aspose.Cells (IsExportComments)
// Description: C# example that creates a workbook, adds a comment to cell A1, enables HtmlSaveOptions.IsExportComments, saves to HTML, reads the file and checks for the "<!--" marker to confirm that comments are exported.
// Keywords: Aspose.Cells HTML export comments | IsExportComments C# | verify HTML comment tag | Aspose.Cells save as HTML | C# workbook comment validation
// Common Searches: Aspose.Cells export comments to HTML | check for <!-- in saved HTML Aspose | C# verify HTML comment tags after export | IsExportComments not working | Aspose.Cells HTMLSaveOptions comment validation
// Developer Intent: Confirm that enabling IsExportComments in HtmlSaveOptions causes the generated HTML file to contain HTML comment delimiters (<!--).
// Use Cases: Automated test to ensure comment export works in HTML output. | Debugging missing comments after saving a workbook as HTML. | Integrating comment verification into a CI pipeline for Aspose.Cells projects.
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to HTML with comments and asserts the presence of "<!--" in the output. | Write a unit test using NUnit that validates HtmlSaveOptions.IsExportComments produces HTML comment tags. | Explain step‑by‑step how to read a saved HTML file and search for comment markers when exporting comments with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentExportValidation
{
    // C# example that creates a workbook, adds a comment to cell A1, enables HtmlSaveOptions.IsExportComments, saves to HTML, reads the file and checks for the "<!--" marker to confirm that comments are exported.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a test comment";

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true   // Enable comment export
            };

            // Save the workbook as HTML
            string htmlPath = "output_with_comments.html";
            workbook.Save(htmlPath, htmlOptions);

            // Read the generated HTML file
            string htmlContent = File.ReadAllText(htmlPath);

            // Validate that the HTML contains comment tags (<!--)
            bool containsCommentTag = htmlContent.Contains("<!--");

            Console.WriteLine($"HTML contains comment tag: {containsCommentTag}");
        }
    }
}
