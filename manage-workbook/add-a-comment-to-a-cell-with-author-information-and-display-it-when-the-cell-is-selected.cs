using Aspose.Cells;
using System;

class AddCommentDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1 (creates the comment if it does not exist)
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];

        // Set the author and the comment text
        comment.Author = "John Doe";
        comment.Note = "Reviewed by John Doe";

        // Make the comment visible when the cell is selected
        comment.IsVisible = true;

        // Save the workbook
        workbook.Save("CommentWithAuthor.xlsx");
    }
}