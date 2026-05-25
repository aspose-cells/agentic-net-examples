using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PrintCommentsToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell A1 and set its text
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a test comment";

        // Configure the worksheet to print comments as they appear on the sheet
        sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

        // Create PDF save options (default settings are sufficient for comments)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file; comments will be rendered according to the PageSetup setting
        workbook.Save("CommentsPrinted.pdf", pdfOptions);
    }
}