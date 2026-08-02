// Title: C# – Compare Aspose.Cells HTML output with DisableDownlevelRevealedComments flag
// Description: Shows how to save an Excel workbook to HTML twice using Aspose.Cells for .NET—once with default settings and once with DisableDownlevelRevealedComments enabled—then reads the files and reports whether the markup differs because of downlevel‑revealed conditional comments.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableDownlevelRevealedComments | downlevel revealed comments | C# Excel to HTML | conditional comments | legacy browser compatibility | HTML markup comparison | Aspose.Cells example | .NET HTML export
// Common Searches: Aspose.Cells disable downlevel revealed comments | HTML output differences with DisableDownlevelRevealedComments | compare generated HTML from Aspose.Cells | how to turn off conditional comments in Aspose.Cells HTML export | C# sample for Aspose.Cells HtmlSaveOptions
// Developer Intent: See how the DisableDownlevelRevealedComments property changes the HTML produced by Aspose.Cells.
// Use Cases: Verify that HTML generated for legacy browsers omits downlevel‑revealed conditional comments. | Create two versions of the same workbook—standard and comment‑free—for cross‑browser testing. | Add an automated check that asserts the markup differs when the flag is toggled.
// AI Prompts: Generate C# code that highlights the exact sections removed when DisableDownlevelRevealedComments is true. | Explain the role of downlevel‑revealed conditional comments in Aspose.Cells HTML export and why a developer might disable them. | Write an NUnit test that saves a workbook with default and disabled options, reads both files, and asserts they are not identical.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to save an Excel workbook to HTML twice using Aspose.Cells for .NET—once with default settings and once with DisableDownlevelRevealedComments enabled—then reads the files and reports whether the markup differs because of downlevel‑revealed conditional comments.
    public class CompareDownlevelRevealedComments
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and add a comment to cell A1
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Comment comment = worksheet.Comments[worksheet.Comments.Add("A1")];
            comment.Note = "Sample comment for downlevel-revealed test";

            // Save HTML with default settings (DisableDownlevelRevealedComments = false)
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            string defaultHtmlPath = "output_default.html";
            workbook.Save(defaultHtmlPath, defaultOptions);

            // Save HTML with DisableDownlevelRevealedComments set to true
            HtmlSaveOptions disabledOptions = new HtmlSaveOptions
            {
                DisableDownlevelRevealedComments = true
            };
            string disabledHtmlPath = "output_disabled.html";
            workbook.Save(disabledHtmlPath, disabledOptions);

            // Load the generated HTML files as plain text, ensuring they exist
            string defaultHtml = File.Exists(defaultHtmlPath)
                ? File.ReadAllText(defaultHtmlPath)
                : string.Empty;

            string disabledHtml = File.Exists(disabledHtmlPath)
                ? File.ReadAllText(disabledHtmlPath)
                : string.Empty;

            // Simple comparison: check if the HTML contents are different
            bool areDifferent = !string.Equals(defaultHtml, disabledHtml, StringComparison.Ordinal);

            Console.WriteLine($"HTML files are {(areDifferent ? "different" : "identical")}.");
            if (areDifferent)
            {
                Console.WriteLine("Differences are due to the handling of downlevel-revealed conditional comments.");
            }
        }
    }
}
