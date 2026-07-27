// Title: Turn off downlevel‑revealed conditional comments in Aspose.Cells HTML export (C#)
// Description: This example creates a workbook, adds a comment to cell A1, sets HtmlSaveOptions.DisableDownlevelRevealedComments to true, and saves the file as HTML. The resulting HTML contains no downlevel‑revealed IE conditional comments.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableDownlevelRevealedComments | C# | .NET | HTML export | conditional comments | IE conditional markup | downlevel revealed | remove IE comments
// Common Searches: Aspose.Cells disable downlevel revealed comments | HtmlSaveOptions property to remove IE conditional comments | Save Excel as HTML without conditional markup in .NET | Turn off downlevel‑revealed comments in Aspose.Cells output
// Developer Intent: Generate HTML from a workbook without embedding downlevel‑revealed IE conditional comments.
// Use Cases: Export Excel reports to clean HTML for modern browsers that do not need IE‑specific markup. | Create HTML email templates from spreadsheets without conditional comments that can break email clients. | Produce standards‑compliant web documentation from Excel files.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML with downlevel‑revealed comments disabled. | Explain how HtmlSaveOptions.DisableDownlevelRevealedComments changes the generated HTML and when it should be applied. | Show a method to verify that the saved HTML file contains no downlevel‑revealed conditional comments.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a comment to cell A1, sets HtmlSaveOptions.DisableDownlevelRevealedComments to true, and saves the file as HTML. The resulting HTML contains no downlevel‑revealed IE conditional comments.
    class DisableDownlevelRevealedCommentsDemo
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to demonstrate that comments exist
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Sample comment";

            // Configure HTML save options to turn off downlevel-revealed conditional comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Setting this property to true disables the downlevel-revealed comments in the output HTML
            htmlOptions.DisableDownlevelRevealedComments = true;

            // Save the workbook as HTML using the configured options (lifecycle: save)
            workbook.Save("OutputWithoutDownlevelComments.html", htmlOptions);

            Console.WriteLine("Workbook saved with downlevel-revealed comments disabled.");
        }
    }
}
