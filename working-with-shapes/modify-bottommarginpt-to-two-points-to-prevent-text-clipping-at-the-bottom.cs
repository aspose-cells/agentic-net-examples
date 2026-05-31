using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);
        shape.Text = "Sample text with increased bottom margin";

        // Set the bottom margin of the shape's text to 2 points
        shape.TextBody.TextAlignment.BottomMarginPt = 2.0;

        // Save the workbook to a file
        workbook.Save("BottomMarginDemo.xlsx");
    }
}