using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ChangeCommentBackground
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can also use workbook.Worksheets["SheetName"])
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell A1 (or retrieve an existing one)
        int commentIdx = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIdx];
        comment.Note = "Comment with custom background";

        // Change the background fill of the comment shape
        comment.CommentShape.FillFormat.IsVisible = true;               // make fill visible
        comment.CommentShape.FillFormat.ForeColor = Color.LightYellow; // background color
        comment.CommentShape.FillFormat.BackColor = Color.Black;       // optional border color
        comment.CommentShape.FillFormat.Transparency = 0.0;            // fully opaque

        // Make the comment visible in the sheet
        comment.IsVisible = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}