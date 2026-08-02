// Title: Export Excel Comments to HTML with Aspose.Cells while Hiding Them in Legacy Browsers
// Description: Creates a workbook, adds a comment to cell A1, marks it invisible, and saves the file as HTML using HtmlSaveOptions (IsExportComments = true, DisableDownlevelRevealedComments = true). The code then reads the output to verify that the comment text is not rendered in older browsers.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableDownlevelRevealedComments | hide comments HTML | legacy browsers | export comments | C# Excel to HTML | conditional comments | workbook to HTML | cell comment visibility
// Common Searches: Aspose.Cells hide Excel comments in HTML for old browsers | HtmlSaveOptions DisableDownlevelRevealedComments example C# | Export comments but keep them invisible in generated HTML | Verify comment text is absent in Aspose.Cells HTML output | C# save workbook as HTML without showing cell comments
// Developer Intent: Generate an HTML representation of an Excel workbook where comments are exported but remain invisible to users of legacy browsers.
// Use Cases: Produce HTML reports for intranet sites that must support IE8 or earlier without displaying cell comments. | Automate a validation step that confirms comment text is omitted from the final HTML markup. | Create archival HTML snapshots that retain comment data for future processing while keeping the UI clean for end‑users.
// AI Prompts: Show how to use Aspose.Cells HtmlSaveOptions to export comments and disable downlevel‑revealed conditional comments for legacy browsers. | Write a C# unit test that asserts the generated HTML does not contain a specific comment when DisableDownlevelRevealedComments is true. | Explain the impact of the DisableDownlevelRevealedComments flag on the HTML markup produced by Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLegacyCommentTest
{
    // Creates a workbook, adds a comment to cell A1, marks it invisible, and saves the file as HTML using HtmlSaveOptions (IsExportComments = true, DisableDownlevelRevealedComments = true). The code then reads the output to verify that the comment text is not rendered in older browsers.
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
            comment.Note = "Legacy browser test comment";

            // Ensure the comment is not visible in the worksheet (optional)
            comment.IsVisible = false;

            // Configure HTML save options:
            // - Export comments so they are written to the HTML.
            // - Disable downlevel‑revealed conditional comments to hide them in legacy browsers.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true,
                DisableDownlevelRevealedComments = true
            };

            // Save the workbook as HTML
            string htmlPath = "LegacyCommentTest.html";
            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"Workbook saved to '{htmlPath}' with comments hidden for legacy browsers.");

            // Simple verification: read the generated HTML and check that the comment text does not appear
            string htmlContent = File.ReadAllText(htmlPath);
            if (htmlContent.Contains("Legacy browser test comment"))
            {
                Console.WriteLine("Warning: Comment text is still present in the HTML output.");
            }
            else
            {
                Console.WriteLine("Success: Comment text is not present in the HTML output.");
            }
        }
    }
}
