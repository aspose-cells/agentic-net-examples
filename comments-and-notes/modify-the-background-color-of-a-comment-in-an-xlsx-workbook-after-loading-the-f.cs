using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ModifyCommentBackground
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is a comment on cell A1; add one if it does not exist
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "Sample comment";

        // Set the background (fill) color of the comment shape
        // ForeColor defines the primary fill color; BackColor can be used for gradient fills
        comment.CommentShape.FillFormat.IsVisible = true;
        comment.CommentShape.FillFormat.ForeColor = Color.LightBlue;   // desired background color
        comment.CommentShape.FillFormat.BackColor = Color.White;      // optional secondary color

        // Optionally make the comment visible when the file is opened
        comment.IsVisible = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}