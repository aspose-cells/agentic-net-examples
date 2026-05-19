using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will act as a comment box
        // Parameters: upper left row, upper left column, row offset, column offset, height, width
        Shape commentBox = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 200);

        // Link the shape to cell G12 so it moves when the cell shifts
        commentBox.LinkedCell = "$G$12";

        // Optionally set some visual properties (e.g., fill color) for clarity
        commentBox.FillFormat.ForeColor = System.Drawing.Color.LightYellow;
        commentBox.LineFormat.Weight = 1.0f;

        // Save the workbook
        workbook.Save("LinkedCommentBox.xlsx");
    }
}