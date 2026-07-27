// Title: Export Workbook to HTML without Comments Using Aspose.Cells for .NET
// Description: Shows how to add a comment to a cell, set HtmlSaveOptions.IsExportComments to false, save the workbook as HTML, and programmatically confirm that the comment markup is not present in the output file.
// Keywords: Aspose.Cells | C# HTML export | IsExportComments false | remove comments from HTML | Aspose.Cells HtmlSaveOptions | verify comment absence | export workbook without comments | .NET spreadsheet export
// Common Searches: Aspose.Cells disable comment export | HTML export without cell comments C# | HtmlSaveOptions IsExportComments usage | check comment removal in generated HTML Aspose.Cells | save workbook as HTML without comments .NET
// Developer Intent: Exclude cell comments from the HTML output when saving a workbook and validate that no comment markup remains.
// Use Cases: Create public‑facing HTML reports that hide internal notes. | Generate compliance‑ready worksheet snapshots without exposing confidential comments. | Automate a post‑save check to ensure comments are omitted from exported HTML files.
// AI Prompts: Write C# code with Aspose.Cells that saves a workbook as HTML while suppressing comments and then verifies the comment text is absent. | Provide a C# unit test that sets HtmlSaveOptions.IsExportComments = false, exports to HTML, and asserts that the generated file contains no comment elements. | Explain how HtmlSaveOptions.IsExportComments works and how to programmatically confirm comment removal after HTML export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportCommentsDemo
{
    // Shows how to add a comment to a cell, set HtmlSaveOptions.IsExportComments to false, save the workbook as HTML, and programmatically confirm that the comment markup is not present in the output file.
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

            // Verify that the comment text does not appear in the generated HTML
            string htmlContent = File.ReadAllText(outputPath);
            bool commentAbsent = !htmlContent.Contains("This is a test comment");

            Console.WriteLine($"Comment absent in HTML: {commentAbsent}");
        }
    }
}
