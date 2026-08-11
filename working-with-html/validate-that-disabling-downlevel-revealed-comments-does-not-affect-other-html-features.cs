// Title: Aspose.Cells .NET: Verify DisableDownlevelRevealedComments Doesn't Break HTML Formatting
// Description: C# sample that creates a workbook, adds a bold cell and an HTML‑styled comment, saves to HTML with DisableDownlevelRevealedComments true and false, then checks that the comment markup and font-weight:bold remain intact, proving the setting doesn't affect other HTML features.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableDownlevelRevealedComments | HTML export .NET | cell comment HTML | bold formatting CSS | C# Aspose.Cells example | HTML validation | downlevel revealed comments | regression testing
// Common Searches: Aspose.Cells DisableDownlevelRevealedComments effect | check HTML export after disabling downlevel comments | C# validate Aspose.Cells HTML output | does DisableDownlevelRevealedComments remove styles | Aspose.Cells comment HTML markup preservation | how to test HtmlSaveOptions in Aspose.Cells
// Developer Intent: Confirm that turning off downlevel‑revealed comments does not alter other HTML output such as cell styles.
// Use Cases: Automated regression test for HTML export when HtmlSaveOptions change | Generate clean HTML reports without conditional comments while keeping formatting | Compare HTML files produced with the flag enabled and disabled for QA | Ensure cell styling persists when comments are hidden in web publishing | Integrate HTML validation into CI/CD pipelines
// AI Prompts: Write a NUnit test that loads the two HTML files saved with DisableDownlevelRevealedComments true/false and asserts that the comment markup and font-weight:bold are present in both. | Create a PowerShell script to batch‑convert Excel workbooks to HTML with downlevel‑revealed comments disabled and log any missing formatting. | Explain how the DisableDownlevelRevealedComments flag changes the generated HTML for cell comments and why other styles remain unaffected. | Generate a GitHub Actions workflow that runs the validation script after each push. | Provide a step‑by‑step guide to troubleshoot missing CSS when using HtmlSaveOptions in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlValidation
{
    // C# sample that creates a workbook, adds a bold cell and an HTML‑styled comment, saves to HTML with DisableDownlevelRevealedComments true and false, then checks that the comment markup and font-weight:bold remain intact, proving the setting doesn't affect other HTML features.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Cell with comment");
            sheet.Cells["A2"].PutValue("Bold text");
            sheet.Cells["A2"].GetStyle().Font.IsBold = true;

            // Add a comment to A1 with HTML content
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.HtmlNote = "<font style='color:#FF0000;'>This is a <b>test</b> comment.</font>";

            // First save: disable downlevel-revealed comments
            HtmlSaveOptions optionsDisable = new HtmlSaveOptions();
            optionsDisable.DisableDownlevelRevealedComments = true;
            string fileWithoutDownlevel = "output_without_downlevel.html";
            workbook.Save(fileWithoutDownlevel, optionsDisable);
            Console.WriteLine($"Saved HTML with DisableDownlevelRevealedComments = true to '{fileWithoutDownlevel}'.");

            // Second save: keep default (downlevel-revealed comments enabled)
            HtmlSaveOptions optionsEnable = new HtmlSaveOptions();
            optionsEnable.DisableDownlevelRevealedComments = false;
            string fileWithDownlevel = "output_with_downlevel.html";
            workbook.Save(fileWithDownlevel, optionsEnable);
            Console.WriteLine($"Saved HTML with DisableDownlevelRevealedComments = false to '{fileWithDownlevel}'.");

            // Load both HTML files as text for simple validation
            string htmlWithout = File.ReadAllText(fileWithoutDownlevel);
            string htmlWith = File.ReadAllText(fileWithDownlevel);

            // Verify that the comment text exists in both outputs
            bool commentInWithout = htmlWithout.Contains("This is a <b>test</b> comment");
            bool commentInWith = htmlWith.Contains("This is a <b>test</b> comment");

            // Verify that the bold formatting of A2 is present (look for 'font-weight:bold')
            bool boldInWithout = htmlWithout.Contains("font-weight:bold");
            bool boldInWith = htmlWith.Contains("font-weight:bold");

            // Output validation results
            Console.WriteLine("\nValidation Results:");
            Console.WriteLine($"Comment present when disabled: {commentInWithout}");
            Console.WriteLine($"Comment present when enabled : {commentInWith}");
            Console.WriteLine($"Bold formatting preserved when disabled: {boldInWithout}");
            Console.WriteLine($"Bold formatting preserved when enabled : {boldInWith}");

            // Simple overall check
            if (commentInWithout && commentInWith && boldInWithout && boldInWith)
                Console.WriteLine("\nDisabling downlevel-revealed comments does not affect other HTML features.");
            else
                Console.WriteLine("\nSome HTML features were affected by the DisableDownlevelRevealedComments setting.");
        }
    }
}
