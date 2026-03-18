using System;
using System.Drawing;
using Aspose.Cells;

class ChangeCommentColor
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the comment at cell A1; add it if it does not exist
        Comment comment = worksheet.Comments["A1"];
        if (comment == null)
        {
            int commentIndex = worksheet.Comments.Add("A1");
            comment = worksheet.Comments[commentIndex];
            comment.Note = "Initial comment text";
        }

        // Change the comment's font color (e.g., to blue)
        comment.Font.Color = Color.Blue;

        // Make the comment visible so the change can be seen in Excel
        comment.IsVisible = true;

        // Save the workbook with the updated comment color
        workbook.Save("output.xlsx");
    }
}