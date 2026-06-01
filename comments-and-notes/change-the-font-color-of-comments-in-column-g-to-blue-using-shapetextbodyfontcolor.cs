using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ChangeCommentFontColorInColumnG
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample comments in column G (index 6)
        for (int row = 0; row < 5; row++)
        {
            // Add a comment to cell G{row+1}
            int commentIndex = worksheet.Comments.Add(row, 6);
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = $"Comment at G{row + 1}";
        }

        // Iterate through all comments and change font color to blue for those in column G
        foreach (Comment comment in worksheet.Comments)
        {
            if (comment.Column == 6) // Column G (zero‑based index)
            {
                // Access the shape associated with the comment
                CommentShape shape = comment.CommentShape;

                // Use the TextBody collection to set the font color.
                // The TextBody collection contains FontSetting objects; we modify the first (and only) one.
                // This changes the font color of the entire comment text.
                shape.TextBody[0].Font.Color = Color.Blue;

                // Ensure the comment is visible (optional)
                comment.IsVisible = true;
            }
        }

        // Save the workbook
        workbook.Save("CommentsColumnG_BlueFont.xlsx");
    }
}