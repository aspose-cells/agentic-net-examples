using System;
using Aspose.Cells;

namespace AsposeCellsCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Sample Cell");

            // Add a comment to the cell if it does not exist
            Comment comment = cell.Comment;
            if (comment == null)
            {
                // Add comment via the worksheet's comment collection
                int commentIndex = worksheet.Comments.Add("A1");
                comment = worksheet.Comments[commentIndex];
            }

            // Set author and note
            comment.Author = "John Doe";
            comment.Note = "This comment appears when the cell is selected.";

            // Make the comment visible (displayed when the cell is selected)
            comment.IsVisible = true;

            // Save the workbook
            workbook.Save("CommentWithAuthor.xlsx");
        }
    }
}