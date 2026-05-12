using System;
using System.Drawing;
using Aspose.Cells;

class ChangeCommentFontColorInColumnG
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Column G has index 6 (zero‑based)
        const int targetColumn = 6;

        // Iterate through all comments in the worksheet
        CommentCollection comments = worksheet.Comments;
        for (int i = 0; i < comments.Count; i++)
        {
            Comment comment = comments[i];

            // Process only comments that belong to column G
            if (comment.Column == targetColumn)
            {
                // Change the font color of the comment text to blue
                comment.Font.Color = Color.Blue;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}