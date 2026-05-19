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

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);

        // Set the shape's text
        shape.Text = "Centered Text";

        // Align text horizontally to center
        shape.TextHorizontalAlignment = TextAlignmentType.Center;

        // Align text vertically to middle (center)
        shape.TextVerticalAlignment = TextAlignmentType.Center;

        // Save the workbook
        workbook.Save("ShapeCenteredAlignment.xlsx");
    }
}