using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyCommentBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is a sample comment";

        // Get the shape associated with the comment
        Shape commentShape = comment.CommentShape;

        // Make sure the shape's fill is visible
        commentShape.IsFilled = true;

        // Apply a solid blue background using the FillFormat's ForeColor property
        commentShape.FillFormat.ForeColor = Color.Blue;

        // Save the workbook
        workbook.Save("CommentWithBlueBackground.xlsx");
    }
}