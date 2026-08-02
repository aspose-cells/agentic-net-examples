using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook and add a sample comment
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "This is a test comment";

        // Configure HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions();
        // Disable downlevel‑revealed conditional comments so older browsers won't see them
        options.DisableDownlevelRevealedComments = true;
        // Do not export comments to HTML (optional, based on requirement)
        options.IsExportComments = false;

        // Save the workbook as HTML with the configured options
        workbook.Save("output.html", options);
    }
}

// Author: Aspose.Cells expert – demonstrates disabling downlevel‑revealed comments when exporting to HTML.