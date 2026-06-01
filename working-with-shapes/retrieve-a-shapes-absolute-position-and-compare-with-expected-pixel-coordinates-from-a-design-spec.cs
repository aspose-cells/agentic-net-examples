using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapePositionCheck
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape.
        // Parameters: upper left row, upper left column, height, width, top offset, left offset
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 80, 100, 0, 0);

        // Set explicit pixel offsets from worksheet borders
        shape.X = 150; // horizontal offset in pixels
        shape.Y = 200; // vertical offset in pixels

        // Retrieve absolute position and size using GetActualBox (returns x, y, w, h)
        float[] actualBox = shape.GetActualBox();
        float actualX = actualBox[0];
        float actualY = actualBox[1];
        float actualWidth = actualBox[2];
        float actualHeight = actualBox[3];

        // Retrieve position using individual properties for verification
        int propX = shape.X;
        int propY = shape.Y;
        int propLeft = shape.Left;
        int propTop = shape.Top;
        int propRight = shape.Right;
        int propBottom = shape.Bottom;

        // Expected design specification (example values)
        int expectedX = 150;
        int expectedY = 200;
        int expectedWidth = 100;
        int expectedHeight = 80;

        // Compare actual values with expected specifications
        bool positionMatches = Math.Abs(actualX - expectedX) < 0.01 && Math.Abs(actualY - expectedY) < 0.01;
        bool sizeMatches = Math.Abs(actualWidth - expectedWidth) < 0.01 && Math.Abs(actualHeight - expectedHeight) < 0.01;

        // Output results
        Console.WriteLine($"Actual Box: X={actualX}, Y={actualY}, Width={actualWidth}, Height={actualHeight}");
        Console.WriteLine($"Properties: X={propX}, Y={propY}, Left={propLeft}, Top={propTop}, Right={propRight}, Bottom={propBottom}");
        Console.WriteLine($"Position matches spec: {positionMatches}");
        Console.WriteLine($"Size matches spec: {sizeMatches}");

        // Save the workbook
        workbook.Save("ShapePositionCheck.xlsx");
    }
}