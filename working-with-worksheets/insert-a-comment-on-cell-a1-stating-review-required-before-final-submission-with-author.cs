using System;
using Aspose.Cells;

class InsertCommentDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];

        // Set the comment text
        comment.Note = "Review required before final submission";

        // Set the comment author
        comment.Author = "John Doe";

        // Save the workbook
        workbook.Save("CommentAdded.xlsx");
    }
}