using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCommentCheck
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a comment to cell A1
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            int commentIdx = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "Standard comment for testing";

            // Configure HTML save options:
            // - Export comments so they appear in the HTML.
            // - Disable downlevel‑revealed conditional comments.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true,
                DisableDownlevelRevealedComments = true
            };

            // Save the workbook as HTML
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the generated HTML does not contain downlevel‑revealed comments
            string htmlContent = File.ReadAllText(htmlPath);
            bool containsDownlevelRevealed = htmlContent.Contains("<![if", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine("Downlevel‑revealed conditional comments present: " + containsDownlevelRevealed);
            Console.WriteLine("HTML contains only standard comment tags: " + (!containsDownlevelRevealed));
        }
    }
}