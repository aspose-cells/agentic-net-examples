using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCommentFontColor
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Access the shape attached to the comment
            CommentShape commentShape = comment.CommentShape;

            // Ensure the shape has a TextBody collection
            // Set the font color of the first text run to red
            // (FontSettingCollection items expose a Font property)
            commentShape.TextBody[0].Font.Color = Color.Red;

            // Optionally make the comment visible
            comment.IsVisible = true;

            // Save the workbook
            workbook.Save("CommentWithRedFont.xlsx");
        }
    }
}