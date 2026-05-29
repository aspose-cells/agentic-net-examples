using System;
using Aspose.Cells;

namespace AsposeCellsLegacyBrowserCommentTest
{
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
            comment.Note = "Legacy browser comment test";

            // Configure HTML save options:
            // - Export comments so they are present in the HTML.
            // - Disable downlevel-revealed conditional comments to hide them in legacy browsers.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                IsExportComments = true,
                DisableDownlevelRevealedComments = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("LegacyBrowserCommentTest.html", saveOptions);

            Console.WriteLine("HTML file saved with comments exported and downlevel-revealed comments disabled.");
        }
    }
}