using System;
using Aspose.Cells;

namespace AsposeCellsCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a comment to cell A1 using the Comments collection
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];

            // Set the comment text and author
            comment.Note = "This is a newly added comment.";
            comment.Author = "DemoUser";

            // Save the workbook with the new comment
            workbook.Save("output.xlsx");
        }
    }
}