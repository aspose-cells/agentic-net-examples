using System;
using Aspose.Cells;

namespace AsposeCellsCommentExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIdx = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "This comment will be visible only in modern browsers";

            // Configure HTML save options:
            // - Export comments (so they appear in modern browsers)
            // - Disable downlevel‑revealed conditional comments, preventing older browsers
            //   from seeing the comment markup.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true,
                DisableDownlevelRevealedComments = true
            };

            // Save the workbook as HTML with the configured options
            workbook.Save("CommentModernOnly.html", htmlOptions);

            Console.WriteLine("Workbook saved as HTML with comments hidden from legacy browsers.");
        }
    }
}