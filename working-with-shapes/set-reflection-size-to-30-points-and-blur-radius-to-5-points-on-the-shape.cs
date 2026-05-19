using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SetReflectionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

        // Access the reflection effect of the shape
        ReflectionEffect reflection = shape.Reflection;

        // Set reflection size to 30 points (percentage of the gradient ramp)
        reflection.Size = 30;

        // Set reflection blur radius to 5 points
        reflection.Blur = 5;

        // Save the workbook with the applied reflection effect
        workbook.Save("ShapeReflectionDemo.xlsx");
    }
}