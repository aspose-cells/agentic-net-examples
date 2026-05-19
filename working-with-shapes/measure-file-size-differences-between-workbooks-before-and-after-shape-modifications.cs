using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeSizeDifferenceDemo
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 50);

        // Save the workbook before modifying the shape (save rule)
        string beforePath = "BeforeShapeModification.xlsx";
        workbook.Save(beforePath);

        // Get the file size before modification
        long sizeBefore = new FileInfo(beforePath).Length;

        // Modify the shape (e.g., change its dimensions)
        rectangle.Width = 200;
        rectangle.Height = 150;

        // Save the workbook after modifying the shape
        string afterPath = "AfterShapeModification.xlsx";
        workbook.Save(afterPath);

        // Get the file size after modification
        long sizeAfter = new FileInfo(afterPath).Length;

        // Output the size comparison
        Console.WriteLine($"File size before modification: {sizeBefore} bytes");
        Console.WriteLine($"File size after modification:  {sizeAfter} bytes");
        Console.WriteLine($"Size difference: {sizeAfter - sizeBefore} bytes");
    }
}