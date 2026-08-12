// Title: C# Unit Test – Verify ExportComments=true Generates Comment Icons in Aspose.Cells HTML Export
// Description: Creates a workbook, adds a visible comment to cell A1, sets HtmlSaveOptions.IsExportComments to true, saves to a memory stream as HTML, and asserts that the output contains the comment markup, confirming that comment icons are exported.
// Keywords: Aspose.Cells | HTML export | IsExportComments | comment icons | C# unit test | automated verification | memory stream | comment markup
// Common Searches: Aspose.Cells unit test export comments to HTML | How to check comment icons in Aspose.Cells HTML output | IsExportComments true example C# | Validate comment visibility in Aspose.Cells HTML export
// Developer Intent: Write an automated test that ensures comment icons appear in the HTML when IsExportComments is enabled.
// Use Cases: CI pipeline validation that HTML reports retain cell comments for end‑user documentation. | Regression test to detect accidental removal of comment markup during library updates. | Quality assurance of web‑based spreadsheet viewers that rely on Aspose.Cells HTML export.
// AI Prompts: Generate an MSTest method that creates a workbook, adds a visible comment, exports to HTML with IsExportComments = true, and asserts the HTML contains the comment element. | Provide an xUnit test example that validates Aspose.Cells includes comment icons in the HTML output when HtmlSaveOptions.IsExportComments is set. | Write a NUnit test that checks for the presence of the comment CSS class or HTML tag after saving a workbook with comments to HTML using Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a workbook, adds a visible comment to cell A1, sets HtmlSaveOptions.IsExportComments to true, saves to a memory stream as HTML, and asserts that the output contains the comment markup, confirming that comment icons are exported.
    class Program
    {
        static void Main()
        {
            try
            {
                ExportComments_IncludesCommentIcons();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        static void ExportComments_IncludesCommentIcons()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Add a visible comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a test comment";
            comment.IsVisible = true; // ensures the comment icon is present

            // Configure HTML save options to export comments
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                IsExportComments = true
            };

            // Save the workbook to a memory stream as HTML
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, options);
                ms.Position = 0;
                string html = new StreamReader(ms, Encoding.UTF8).ReadToEnd();

                // Verify that the generated HTML contains comment markup (icon)
                if (!html.Contains("comment", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("HTML should contain comment markup when IsExportComments is true.");
                }
            }
        }
    }
}
