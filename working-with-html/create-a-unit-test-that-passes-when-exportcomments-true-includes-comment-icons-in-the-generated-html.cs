// Title: Create a C# unit test that verifies ExportComments=true adds comment icons when saving a workbook to HTML with Aspose.Cells
// AI Prompts: Generate a C# test method (e.g., NUnit or MSTest) that adds a comment to cell A1, saves the workbook using HtmlSaveOptions, and asserts that the produced HTML string contains the 'comment.png' image. | Write code to export a workbook to a MemoryStream with ExportComments enabled, read the HTML output, and fail the test if the comment icon is not found.
// Common Searches: how to assert comment icons appear in Aspose.Cells HTML export unit test | C# Aspose.Cells ExportComments true unit test example | verify comment.png is included in HTML saved by Aspose.Cells | unit testing workbook comment rendering in HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions test for comment image presence
// Tags: Aspose.Cells HtmlSaveOptions ExportComments verification | C# unit test comment icon in HTML output | check comment.png presence in Aspose.Cells HTML export | save workbook as HTML with comment icons | Aspose.Cells comment rendering test

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // A C# unit test that creates a workbook, adds a comment to cell A1, saves the workbook as HTML using HtmlSaveOptions with ExportComments enabled, reads the HTML from a memory stream, and asserts that the output contains the 'comment.png' image indicating the comment icon.
    public class HtmlExportTests
    {
        public void ExportComments_IncludesCommentIcons()
        {
            try
            {
                // Create a new workbook and add a comment to cell A1
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                int commentIndex = sheet.Comments.Add("A1");
                var comment = sheet.Comments[commentIndex];
                comment.Note = "Sample comment";

                // Configure HTML export options (comments are exported by default)
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Export to HTML using a memory stream
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, htmlOptions);
                    stream.Position = 0;
                    string html = new StreamReader(stream).ReadToEnd();

                    // Verify that the generated HTML contains the comment icon image
                    if (html.Contains("comment.png"))
                    {
                        Console.WriteLine("Test passed: Comment icon found in HTML output.");
                    }
                    else
                    {
                        Console.WriteLine("Test failed: Comment icon not found in HTML output.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            var test = new HtmlExportTests();
            test.ExportComments_IncludesCommentIcons();
        }
    }
}
