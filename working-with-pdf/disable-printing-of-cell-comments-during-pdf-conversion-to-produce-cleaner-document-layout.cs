using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Add a comment to demonstrate that it will not be printed
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This comment will not appear in the PDF";

        // Disable printing of comments when converting to PDF
        sheet.PageSetup.PrintComments = PrintCommentsType.PrintNoComments;

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}

// Author: Aspose.Cells .NET example – disables comment printing for PDF output.