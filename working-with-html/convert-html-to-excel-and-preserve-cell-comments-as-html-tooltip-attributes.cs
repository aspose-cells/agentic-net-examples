using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Sample");
        worksheet.Cells["B1"].PutValue("Data");

        // Add a comment to cell A1
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "This is a tooltip comment";

        // Set HTML save options to export comments as tooltip attributes
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.IsExportComments = true;                     // Enable comment export
        htmlOptions.ExportCommentsType = PrintCommentsType.PrintInPlace; // Export as tooltip (title attribute)

        // Save the workbook as HTML with the configured options
        workbook.Save("output.html", htmlOptions);
    }
}