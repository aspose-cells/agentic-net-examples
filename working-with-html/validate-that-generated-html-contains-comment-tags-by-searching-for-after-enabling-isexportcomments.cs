// Title: Aspose.Cells C# – Validate HTML Export Includes <!-- Comment Tags (IsExportComments=True)
// Description: This C# example creates a workbook, adds a comment to cell A1, saves it as HTML with HtmlSaveOptions.IsExportComments enabled, reads the generated file, and confirms the presence of the HTML comment delimiter <!--. It shows how to programmatically verify comment export in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | HTML export comments | IsExportComments | HtmlSaveOptions | validate HTML comment tags | export cell comments to HTML | Aspose.Cells .NET example | HTML comment delimiter | automated test Aspose.Cells
// Common Searches: Aspose.Cells export comments to HTML | IsExportComments true example | C# check for <!-- in exported HTML | validate Aspose.Cells HTML output | how to include cell comments in HTML with Aspose.Cells
// Developer Intent: Confirm that enabling IsExportComments in HtmlSaveOptions causes Aspose.Cells to embed cell comments as HTML <!-- comment tags> in the saved file.
// Use Cases: Generate documentation‑ready HTML reports that retain Excel cell comments. | Create automated regression tests that ensure comment export works after library updates. | Batch‑process multiple workbooks to HTML while verifying each output contains comment markers. | Integrate comment‑preserving HTML export into a web service that serves Excel data as web‑friendly pages.
// AI Prompts: Write a C# function that accepts an Excel file path, saves it as HTML with comments exported, and returns true if the output contains the <!-- marker. | Provide a unit test using NUnit that loads a workbook, exports to HTML with IsExportComments=true, and asserts the presence of comment tags. | Generate a PowerShell script that leverages Aspose.Cells .NET to convert Excel files to HTML with comments and logs validation results. | Explain how to configure HtmlSaveOptions to include cell comments and how to programmatically verify the export in a CI pipeline.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentExportValidation
{
    // This C# example creates a workbook, adds a comment to cell A1, saves it as HTML with HtmlSaveOptions.IsExportComments enabled, reads the generated file, and confirms the presence of the HTML comment delimiter <!--. It shows how to programmatically verify comment export in Aspose.Cells for .NET.
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
                IsExportComments = true // Enable comment export
            };

            // Define output HTML file path
            string htmlPath = "output_with_comments.html";

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, htmlOptions);

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Validate that the HTML contains comment tags (<!--)
            bool containsCommentTag = htmlContent.Contains("<!--");

            // Output validation result
            Console.WriteLine($"HTML contains comment tags: {containsCommentTag}");
        }
    }
}
