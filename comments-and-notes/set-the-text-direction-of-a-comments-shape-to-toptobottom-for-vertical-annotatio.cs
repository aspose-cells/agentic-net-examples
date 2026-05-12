using System;
using Aspose.Cells;

namespace AsposeCellsCommentOrientation
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
            comment.Note = "This comment will be displayed vertically.";

            // Set the comment's text orientation to TopToBottom (stacked vertical)
            comment.TextOrientationType = TextOrientationType.TopToBottom;

            // Save the workbook
            workbook.Save("CommentTopToBottom.xlsx");
        }
    }
}