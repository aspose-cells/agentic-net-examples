using System;
using Aspose.Cells;

namespace AsposeCellsCommentOrientation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This comment will be displayed vertically (top‑to‑bottom).";

            // Set the text orientation of the comment's shape to TopToBottom
            // This changes the layout of the comment to stacked vertical text.
            comment.TextOrientationType = TextOrientationType.TopToBottom;

            // Save the workbook (lifecycle: save)
            workbook.Save("CommentTopToBottom.xlsx");
        }
    }
}