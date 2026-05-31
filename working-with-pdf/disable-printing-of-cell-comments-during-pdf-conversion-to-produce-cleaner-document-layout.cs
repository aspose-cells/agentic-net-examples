using System;
using Aspose.Cells;

class DisableCommentsPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add a comment to demonstrate that it will not be printed
        int commentIndex = worksheet.Comments.Add("A1");
        worksheet.Comments[commentIndex].Note = "This comment will not appear in the PDF.";

        // Disable printing of comments for this worksheet
        worksheet.PageSetup.PrintComments = PrintCommentsType.PrintNoComments;

        // Save the workbook as a PDF; comments are omitted from the output
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}