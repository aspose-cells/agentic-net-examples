using System;
using System.Drawing;
using Aspose.Cells;

class ChangeCommentFontColor
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Try to get an existing comment on cell A1
        Comment comment = worksheet.Comments["A1"];

        // If the comment does not exist, add a new one
        if (comment == null)
        {
            int commentIndex = worksheet.Comments.Add("A1");
            comment = worksheet.Comments[commentIndex];
            comment.Note = "Sample comment";
        }

        // Change the font color of the comment (e.g., to Blue)
        comment.Font.Color = Color.Blue;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}