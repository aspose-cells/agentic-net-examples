using System;
using Aspose.Cells;

class AddCommentExample
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");

        // Get the first worksheet (you can change the index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell C3 (row index 2, column index 2)
        // Using the overload that accepts a cell name
        int commentIndex = worksheet.Comments.Add("C3");

        // Retrieve the newly added comment object
        Comment comment = worksheet.Comments[commentIndex];

        // Set comment properties
        comment.Note = "This is a comment added via Aspose.Cells.";
        comment.Author = "AsposeUser";
        comment.IsVisible = true; // Make the comment visible by default

        // Save the workbook with the new comment
        workbook.Save("Output.xlsx");
    }
}