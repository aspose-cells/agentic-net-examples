using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – demonstrates disabling downlevel‑revealed conditional comments in HTML export
    class HtmlSaveOptionsDisableDownlevelRevealedCommentsDemo
    {
        static void Main()
        {
            // Create a new workbook and add a comment (optional, just to illustrate output)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Sample comment";

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Turn off downlevel‑revealed conditional comments (default is false)
                DisableDownlevelRevealedComments = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with DisableDownlevelRevealedComments = true");
        }
    }
}