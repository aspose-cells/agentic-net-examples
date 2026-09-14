// Title: Generate HTML from an Aspose.Cells workbook in C# while hiding cell comments for legacy browser support
// AI Prompts: Write a C# console program that creates a workbook, adds a comment to a cell, saves it as HTML using Aspose.Cells with comments excluded, and displays the output file path. | Configure Aspose.Cells HtmlSaveOptions to disable comment export and show how to confirm that the comment text is not present in the generated HTML. | Add code that reads the saved HTML file, searches for a specific comment string, and prints a verification result indicating whether the comment was successfully hidden.
// Common Searches: how to export Excel to HTML with Aspose.Cells without comments for IE8 | C# Aspose.Cells HtmlSaveOptions hide cell comments legacy browsers | verify that Aspose.Cells generated HTML does not contain cell comments | prevent comment export when saving workbook as HTML using Aspose.Cells C# | legacy browser compatible HTML output from Aspose.Cells workbook
// Tags: Aspose.Cells HtmlSaveOptions hide comments | C# export workbook to HTML without comments | legacy browser compatible HTML from Excel | cell comment exclusion Aspose.Cells | verify HTML comment omission C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLegacyHtmlTest
{
    // The example creates a workbook, adds sample data and a comment to cell A1, then saves the workbook as an HTML file using HtmlSaveOptions that omit comments. After saving, it reads the HTML file to ensure the comment text is not present, printing the file location and verification result.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the output HTML file
                string htmlPath = "LegacyCommentTest.html";

                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Add a comment to cell A1
                int commentIndex = sheet.Comments.Add("A1");
                Comment comment = sheet.Comments[commentIndex];
                comment.Note = "This is a test comment that should be hidden in legacy browsers.";

                // Configure HTML save options (comments are not exported by default)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true
                    // Additional options can be set here if needed for legacy compatibility
                };

                // Save the workbook as an HTML file using the configured options
                workbook.Save(htmlPath, saveOptions);

                // Verify that the comment text does not appear in the generated HTML
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    bool commentFound = htmlContent.Contains("This is a test comment");

                    Console.WriteLine($"HTML file generated at: {Path.GetFullPath(htmlPath)}");
                    Console.WriteLine($"Comment hidden in HTML: {!commentFound}");

                    // Optional: Output a snippet of the HTML for manual inspection
                    Console.WriteLine("\n--- HTML snippet ---");
                    Console.WriteLine(htmlContent.Substring(0, Math.Min(500, htmlContent.Length)));
                }
                else
                {
                    Console.WriteLine($"Failed to generate HTML file at: {Path.GetFullPath(htmlPath)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
