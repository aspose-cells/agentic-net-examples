using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell B2
        int commentIndex = sheet.Comments.Add("B2");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "This is the comment text";

        // Retrieve the shape associated with the comment
        CommentShape commentShape = comment.CommentShape;

        // Optional: adjust shape size and appearance
        commentShape.Width = 200;
        commentShape.Height = 100;
        commentShape.FillFormat.ForeColor = Color.LightYellow;
        commentShape.LineFormat.ForeColor = Color.Blue;

        // Set alternative text – this appears as a tooltip when the mouse hovers over the comment
        commentShape.AlternativeText = "Tooltip: additional information";

        // Save the workbook
        workbook.Save("CommentWithTooltip.xlsx");
    }
}