using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeZOrderDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add two overlapping rectangle shapes
        Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
        Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

        // Capture initial Z-order positions
        int initialPos1 = shape1.ZOrderPosition;
        int initialPos2 = shape2.ZOrderPosition;

        // Bring shape2 to the front (positive order)
        shape2.ToFrontOrBack(1);
        int frontPos = shape2.ZOrderPosition;

        // Send shape2 to the back (negative order)
        shape2.ToFrontOrBack(-1);
        int backPos = shape2.ZOrderPosition;

        // Output comparison results
        Console.WriteLine($"Initial ZOrder - Shape1: {initialPos1}, Shape2: {initialPos2}");
        Console.WriteLine($"After ToFrontOrBack(1) - Shape2 ZOrder: {frontPos}");
        Console.WriteLine($"After ToFrontOrBack(-1) - Shape2 ZOrder: {backPos}");
        Console.WriteLine($"Moved to front: {frontPos > initialPos2}");
        Console.WriteLine($"Moved to back: {backPos < frontPos}");

        // Save the workbook
        workbook.Save("ShapeZOrderDemo.xlsx");
    }
}