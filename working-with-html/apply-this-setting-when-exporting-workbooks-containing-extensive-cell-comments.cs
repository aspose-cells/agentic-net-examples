using System;
using Aspose.Cells;

class ExportCommentsExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a large number of comments to simulate an extensive comment scenario
        // Here we add 1000 comments distributed across the first 10 rows and columns
        for (int i = 0; i < 1000; i++)
        {
            int row = i / 10;          // Row index (0‑based)
            int col = i % 10;          // Column index (0‑based)
            int commentIndex = worksheet.Comments.Add(row, col);
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = $"Comment #{i + 1}";
        }

        // Configure HTML save options to ensure comments are exported
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export comments when saving to HTML
            IsExportComments = true,

            // Choose how comments appear in the HTML output (in‑place display)
            ExportCommentsType = PrintCommentsType.PrintInPlace,

            // Disable downlevel‑revealed conditional comments for better performance
            DisableDownlevelRevealedComments = true
        };

        // Save the workbook to an HTML file using the configured options
        workbook.Save("WorkbookWithExtensiveComments.html", htmlOptions);
    }
}