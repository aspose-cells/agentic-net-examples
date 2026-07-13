using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example demonstrating that disabling downlevel-revealed comments
    // does not interfere with other HTML export features.
    class DisableDownlevelRevealedCommentsDemo
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Sample comment for HTML export";

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions();

            // Disable downlevel-revealed conditional comments
            options.DisableDownlevelRevealedComments = true;

            // Keep other HTML features at their default values (or explicitly set them)
            options.IsExportComments = true;          // Export comments (should still work)
            options.DisableCss = false;               // Use external CSS (default)
            options.IsJsBrowserCompatible = true;    // JavaScript compatibility (default)

            // Save the workbook as HTML using the configured options
            string outputPath = "output.html";
            workbook.Save(outputPath, options);

            Console.WriteLine($"Workbook saved to '{outputPath}' with DisableDownlevelRevealedComments = true.");
        }
    }
}