using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell A1 with custom formatting
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a formatted comment.";
        comment.Font.Name = "Calibri";
        comment.Font.Size = 11;
        comment.Font.Color = Color.Blue;
        comment.IsVisible = true; // make the comment visible on the sheet

        // Configure the sheet to print comments as they appear on the sheet
        sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

        // Save the workbook to PDF; comment formatting is preserved
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("CommentsPrinted.pdf", pdfOptions);
    }
}

// Author: Example demonstrating how to enable printing of cell comments while saving to PDF using Aspose.Cells.