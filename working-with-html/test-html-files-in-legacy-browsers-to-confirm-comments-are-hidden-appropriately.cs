// Title: Export Excel Cell Comments to HTML While Suppressing Them in Legacy Browsers – Aspose.Cells for .NET
// Description: C# sample that adds a comment to cell A1, then saves the workbook as HTML with HtmlSaveOptions.IsExportComments = true and DisableDownlevelRevealedComments = true, preventing the comment from appearing in old browsers like IE6/7.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportComments | DisableDownlevelRevealedComments | legacy browsers | IE6 | HTML export | Excel comments | conditional comments
// Common Searches: Aspose.Cells hide Excel comments in old browsers | HtmlSaveOptions DisableDownlevelRevealedComments example | export workbook to HTML without conditional comments | C# generate HTML from Excel with hidden comments | test comment visibility in legacy browsers Aspose
// Developer Intent: Generate HTML from an Excel file where cell comments are stored but not rendered in outdated browsers.
// Use Cases: Create web‑ready reports that keep internal notes invisible to IE6/7 users. | Automate CI checks to verify that exported HTML lacks downlevel‑revealed comment markup. | Produce documentation where comments are only visible in modern browsers supporting standard conditional comments.
// AI Prompts: Modify the code to also hide comments in current browsers using CSS. | Write a unit test that asserts the HTML output contains no downlevel‑revealed comment tags when DisableDownlevelRevealedComments is true. | Explain the difference between downlevel‑revealed and downlevel‑hidden conditional comments and how Aspose.Cells processes each.

using System;
using Aspose.Cells;

// C# sample that adds a comment to cell A1, then saves the workbook as HTML with HtmlSaveOptions.IsExportComments = true and DisableDownlevelRevealedComments = true, preventing the comment from appearing in old browsers like IE6/7.
class TestLegacyBrowserComments
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
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export comments to the HTML output
            IsExportComments = true,
            // Disable downlevel‑revealed conditional comments so they are not shown in legacy browsers
            DisableDownlevelRevealedComments = true
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("LegacyBrowserComment.html", htmlOptions);

        Console.WriteLine("HTML file saved with comments hidden for legacy browsers.");
    }
}
