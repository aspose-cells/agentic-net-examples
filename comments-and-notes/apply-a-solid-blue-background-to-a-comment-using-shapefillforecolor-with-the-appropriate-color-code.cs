using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentBackgroundDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This comment has a solid blue background.";

            // Access the shape attached to the comment
            Shape commentShape = comment.CommentShape;

            // Ensure the fill format is visible
            commentShape.FillFormat.IsVisible = true;

            // Apply a solid blue background using FillFormat.ForeColor
            commentShape.FillFormat.ForeColor = Color.Blue;

            // Save the workbook (save rule)
            workbook.Save("CommentBlueBackground.xlsx");
        }
    }
}