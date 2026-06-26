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
            comment.Note = "Sample comment";

            // Obtain the shape that represents the comment
            CommentShape commentShape = comment.CommentShape;

            // The TextBody property returns a collection of FontSetting objects.
            // Each FontSetting contains a Font that can be formatted.
            // Here we modify the first (and only) FontSetting to set the font color to red.
            commentShape.TextBody[0].Font.Color = Color.Red;

            // Optionally make the comment visible to see the effect
            comment.IsVisible = true;

            // Save the workbook
            workbook.Save("CommentWithRedFont.xlsx");
        }
    }
}