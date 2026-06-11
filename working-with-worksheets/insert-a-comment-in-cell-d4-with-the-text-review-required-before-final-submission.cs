using System;
using Aspose.Cells;

class InsertCommentDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell D4
        int commentIndex = worksheet.Comments.Add("D4");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "Review required before final submission.";

        // Save the workbook
        workbook.Save("CommentInD4.xlsx");
    }
}