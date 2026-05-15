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

        // Add a comment to cell B2 (using cell name)
        int commentIndex = worksheet.Comments.Add("B2");
        Comment comment = worksheet.Comments[commentIndex];

        // Set author information and comment text
        comment.Author = "John Doe";
        comment.Note = "Reviewed by John Doe.";

        // Make the comment visible when the cell is selected
        comment.IsVisible = true;

        // Save the workbook
        workbook.Save("CommentWithAuthor.xlsx");
    }
}