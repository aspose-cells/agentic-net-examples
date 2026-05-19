using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AdjustShapeZOrder
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add three overlapping shapes to establish a Z-order stack
        Shape shape1 = sheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
        Shape shape2 = sheet.Shapes.AddRectangle(20, 20, 100, 100, 0, 0);
        Shape shape3 = sheet.Shapes.AddRectangle(35, 35, 100, 100, 0, 0);

        // Set explicit Z-order positions (optional, for clarity)
        shape1.ZOrderPosition = 0; // backmost
        shape2.ZOrderPosition = 1;
        shape3.ZOrderPosition = 2; // frontmost

        // Choose the shape whose Z-order will be adjusted
        Shape targetShape = shape2;

        // Retrieve the current Z-order position
        int currentZ = targetShape.ZOrderPosition;
        Console.WriteLine("Current Z-order position: " + currentZ);

        // Calculate the new Z-order by subtracting three
        int newZ = currentZ - 3;

        // Ensure the new Z-order is not negative (minimum is 0)
        if (newZ < 0) newZ = 0;

        // Apply the new Z-order position
        targetShape.ZOrderPosition = newZ;

        // Verify and output the updated Z-order position
        Console.WriteLine("Updated Z-order position: " + targetShape.ZOrderPosition);

        // Save the workbook to persist changes
        workbook.Save("AdjustedZOrder.xlsx");
    }
}