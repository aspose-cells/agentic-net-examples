using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeFrontBackDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add two rectangle shapes to the worksheet
        Shape shape1 = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
        Shape shape2 = worksheet.Shapes.AddRectangle(20, 20, 100, 100, 0, 0);

        // Bring shape2 to the front (positive value)
        shape2.ToFrontOrBack(1);

        // Immediately send shape2 to the back (negative value)
        shape2.ToFrontOrBack(-1);

        // Verify the final Z-order positions of both shapes
        Console.WriteLine("Shape1 ZOrderPosition: " + shape1.ZOrderPosition);
        Console.WriteLine("Shape2 ZOrderPosition: " + shape2.ZOrderPosition);

        // Save the workbook to a file
        workbook.Save("ShapeFrontBackDemo.xlsx");
    }
}