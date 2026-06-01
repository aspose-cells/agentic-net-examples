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

        // Add two rectangle shapes to the worksheet
        // Parameters: upper left row, upper left column, left offset, top offset, width, height
        Shape shape1 = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 100);
        Shape shape2 = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 150, 150);

        // Set absolute pixel positions for the shapes
        shape1.Left = 50;   // horizontal offset from left column (pixels)
        shape1.Top = 80;    // vertical offset from top row (pixels)

        shape2.Left = 200;
        shape2.Top = 300;

        // Calculate horizontal and vertical pixel differences
        int deltaX = shape2.Left - shape1.Left;
        int deltaY = shape2.Top - shape1.Top;

        // Compute Euclidean distance (pixel offset) between the two shapes
        double pixelOffset = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);

        // Store the calculated offset in cell A1
        worksheet.Cells["A1"].PutValue(pixelOffset);

        // Save the workbook to a file
        workbook.Save("ShapeOffsetResult.xlsx");
    }
}