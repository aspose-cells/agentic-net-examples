using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape (parameters: upper left row, column, width, height, left offset, top offset)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 100, 200, 0, 0);

        // Set the rotation angle to 90 degrees for diagonal orientation
        shape.RotationAngle = 90;

        // Optional: add some text to visualize the rotation
        shape.Text = "Diagonal";

        // Save the workbook
        workbook.Save("ShapeDiagonal.xlsx");
    }
}