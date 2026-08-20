// Title: Export Excel to HTML without comments using Aspose.Cells for .NET
// Description: Creates a workbook, adds a value and a comment to cell A1, sets HtmlSaveOptions.IsExportComments to false, saves as HTML, reads the file, and confirms the comment text is absent.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | IsExportComments | export HTML without comments | verify comment removal | Excel to HTML conversion | Aspose.Cells example
// Common Searches: Aspose.Cells disable comment export HTML | How to hide comments when saving Excel as HTML in .NET | Check that comments are not in generated HTML Aspose.Cells | C# export workbook to HTML without comments
// Developer Intent: Prevent comments from being written to the HTML file when converting an Excel workbook and programmatically verify their absence.
// Use Cases: Produce clean HTML reports that omit internal worksheet comments. | Publish Excel‑derived web pages without exposing confidential notes. | Automated testing to ensure comment data is not leaked in HTML exports.
// AI Prompts: Show C# code that saves an Aspose.Cells workbook to HTML with comments excluded and validates the result. | Explain the effect of HtmlSaveOptions.IsExportComments on the generated HTML and how to detect comment remnants. | Create an MSTest unit test that asserts no comment text appears in the saved HTML file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportCommentsDemo
{
    // Creates a workbook, adds a value and a comment to cell A1, sets HtmlSaveOptions.IsExportComments to false, saves as HTML, reads the file, and confirms the comment text is absent.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to a cell
            sheet.Cells["A1"].PutValue("Sample Data");

            // Add a comment to the same cell
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a test comment";

            // Configure HTML save options to NOT export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = false // Ensure comments are excluded
            };

            // Define output HTML file path
            string outputPath = "output_without_comments.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            // Read the generated HTML file
            string htmlContent = File.ReadAllText(outputPath);

            // Verify that the comment text is not present in the HTML
            bool commentAbsent = !htmlContent.Contains("This is a test comment");

            Console.WriteLine($"Comment absent from HTML: {commentAbsent}");
        }
    }
}
