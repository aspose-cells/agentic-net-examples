// Title: Aspose.Cells for .NET: Export Excel Comments to HTML with Modern‑Browser‑Only Visibility
// Description: Learn how to save an Excel workbook as HTML using Aspose.Cells while exporting cell comments that are displayed only in browsers supporting modern comment standards. The example disables down‑level revealed conditional comments and enables comment export, ensuring legacy browsers do not see the notes.
// Keywords: Aspose.Cells | HtmlSaveOptions | IsExportComments | DisableDownlevelRevealedComments | C# export Excel to HTML | hide comments legacy browsers | modern browser comment visibility | Excel to HTML sample | GitHub Aspose.Cells example | code snippet for HTML export
// Common Searches: Aspose.Cells hide Excel comments from old browsers | C# export workbook to HTML without conditional comments | DisableDownlevelRevealedComments usage | IsExportComments true Aspose.Cells | Export Excel notes to HTML for modern browsers only
// Developer Intent: Generate an HTML file from an Excel workbook where cell comments are included for browsers that understand modern comment syntax but are suppressed for legacy browsers.
// Use Cases: Creating web‑ready financial reports that show cell notes only to up‑to‑date browsers. | Building intranet dashboards where legacy IE versions must not display Excel comments. | Publishing documentation with contextual hints that are invisible to browsers lacking modern comment support.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as HTML, enabling comment export while disabling down‑level revealed conditional comments. | Explain the interaction between DisableDownlevelRevealedComments and IsExportComments in Aspose.Cells HTML conversion. | Show how to add multiple cell comments and export them so they appear only in browsers that support modern HTML comment standards.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentExportDemo
{
    // Learn how to save an Excel workbook as HTML using Aspose.Cells while exporting cell comments that are displayed only in browsers supporting modern comment standards. The example disables down‑level revealed conditional comments and enables comment export, ensuring legacy browsers do not see the notes.
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
            comment.Note = "This comment will be exported only to modern browsers.";

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Disable downlevel-revealed conditional comments so older browsers won't see the comment
                DisableDownlevelRevealedComments = true,

                // Export comments (so modern browsers receive them)
                IsExportComments = true
            };

            // Save the workbook as HTML with the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with comments hidden from non‑modern browsers.");
        }
    }
}
