using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCommentFontColorDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];

            // Set the comment text
            comment.Note = "This comment has a custom font color.";

            // Modify the comment's font color (access the Font property of the comment)
            comment.Font.Color = Color.Blue; // Desired color

            // Optionally make the comment visible
            comment.IsVisible = true;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CommentWithCustomFontColor.xlsx");
        }
    }
}