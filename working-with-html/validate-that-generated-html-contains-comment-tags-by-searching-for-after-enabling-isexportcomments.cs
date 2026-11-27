// Title: Export Excel cell comments to HTML and verify that <!-- comment tags are present using Aspose.Cells for .NET (C#)
// Common Searches: how to export a worksheet comment to HTML with Aspose.Cells in C# | C# code to check if exported HTML contains <!-- comment markers from Aspose.Cells | Aspose.Cells HtmlSaveOptions example for saving comments in HTML format | read generated HTML file and search for comment tags using .NET File API | verify that cell comments are included when saving workbook as HTML with Aspose.Cells
// Tags: Aspose.Cells export comments to HTML | HtmlSaveOptions SaveFormat.Html usage | C# verify HTML comment markers | read HTML file with File.ReadAllText .NET | add visible cell comment Aspose.Cells worksheet | check exported HTML contains <!-- tag

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // Required for HtmlSaveOptions

namespace AsposeCellsCommentExportCheck
{
    // The sample creates a workbook, adds a visible comment to cell A1, saves the workbook as HTML using HtmlSaveOptions, reads the resulting HTML file, and confirms that the output includes an HTML comment marker ("<!--"). It prints a success or failure message based on the presence of the comment tag.
    class Program
    {
        static void Main(string[] args)
        {
            const string htmlOutputPath = "CommentExportTest.html";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Add a comment to cell A1
                int commentIndex = sheet.Comments.Add("A1");
                var comment = sheet.Comments[commentIndex];
                comment.Note = "This is a test comment that should appear in the exported HTML.";
                comment.IsVisible = true; // Ensure the comment is visible

                // 3. Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // In recent versions ExportComments property is not required; comments are exported by default.
                // If a specific property exists in your version, set it here.

                // 4. Save the workbook as HTML using the configured options
                workbook.Save(htmlOutputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to HTML at: {Path.GetFullPath(htmlOutputPath)}");

                // 5. Verify that the generated HTML file exists before reading
                if (!File.Exists(htmlOutputPath))
                {
                    Console.WriteLine("ERROR: HTML file was not created.");
                    return;
                }

                // 6. Load the generated HTML file and verify that comment tags are present
                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(htmlOutputPath);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine("Failed to read the generated HTML file:");
                    Console.WriteLine(readEx.Message);
                    return;
                }

                bool containsCommentTag = htmlContent.Contains("<!--");

                // 7. Output the verification result
                if (containsCommentTag)
                {
                    Console.WriteLine("SUCCESS: The exported HTML contains comment tags (\"<!--\").");
                }
                else
                {
                    Console.WriteLine("FAILURE: The exported HTML does NOT contain comment tags (\"<!--\").");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
