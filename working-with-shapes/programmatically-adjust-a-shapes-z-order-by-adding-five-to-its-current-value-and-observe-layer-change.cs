using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AdjustShapeZOrder
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two overlapping shapes
        Shape shape1 = sheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
        Shape shape2 = sheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

        // Display initial Z-order position of shape2
        Console.WriteLine("Initial ZOrderPosition of shape2: " + shape2.ZOrderPosition);

        // Increase Z-order of shape2 by 5 positions
        int newZOrder = shape2.ZOrderPosition + 5;
        shape2.ZOrderPosition = newZOrder;

        // Display new Z-order position to confirm change
        Console.WriteLine("New ZOrderPosition of shape2: " + shape2.ZOrderPosition);

        // Save the workbook to observe the layer change
        workbook.Save("ShapeZOrderAdjusted.xlsx");
    }
}