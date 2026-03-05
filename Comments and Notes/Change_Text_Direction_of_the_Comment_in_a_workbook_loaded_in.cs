using Aspose.Cells;
using System;

class ChangeCommentTextDirection
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Try to get a comment at cell A1; if it does not exist, add one
        Comment comment = worksheet.Comments["A1"];
        if (comment == null)
        {
            worksheet.Comments.Add("A1");
            comment = worksheet.Comments["A1"];
        }

        // Set the comment text (optional, just to have visible content)
        comment.Note = "This comment's text direction has been changed.";

        // Change the text orientation of the comment.
        // Example: rotate text 90 degrees clockwise.
        comment.TextOrientationType = TextOrientationType.ClockWise;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}