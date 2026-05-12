using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SetCommentShapeTextDirection
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIndex];
        comment.Note = "مثال على النص من اليمين إلى اليسار"; // Sample Arabic text

        // Retrieve the shape associated with the comment
        CommentShape shape = comment.CommentShape;

        // Set the text direction of the comment's shape to RightToLeft
        shape.TextDirection = TextDirectionType.RightToLeft;

        // Save the workbook
        workbook.Save("CommentShapeRightToLeft.xlsx");
    }
}