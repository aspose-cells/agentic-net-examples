using System;
using Aspose.Cells;

// Author: Generated example demonstrating how to export comments with
// downlevel‑revealed conditional comments so they are hidden in legacy browsers.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "This comment should be hidden in legacy browsers";

        // Configure HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            // Export comments so they are written to the HTML output
            IsExportComments = true,

            // Keep downlevel‑revealed conditional comments (default is false, which disables them)
            // Setting this to false generates the conditional comments that legacy browsers will hide.
            DisableDownlevelRevealedComments = false
        };

        // Save the workbook as HTML with the specified options
        workbook.Save("LegacyBrowserCommentTest.html", options);

        Console.WriteLine("HTML file saved. Open it in a legacy browser to verify that the comment is hidden.");
    }
}