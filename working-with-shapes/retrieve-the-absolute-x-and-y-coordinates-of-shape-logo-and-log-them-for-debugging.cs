using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RetrieveShapeCoordinates
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape and name it "Logo"
        // Parameters: upperLeftRow, upperLeftColumn, top (pixels), left (pixels), height (pixels), width (pixels)
        Shape logoShape = worksheet.Shapes.AddRectangle(2, 2, 150, 120, 80, 200);
        logoShape.Name = "Logo";

        // Optionally set explicit position (pixels from worksheet top‑left)
        logoShape.X = 300; // horizontal offset
        logoShape.Y = 200; // vertical offset

        // Retrieve the absolute X and Y coordinates of the shape
        int absoluteX = logoShape.X; // pixels from worksheet left border
        int absoluteY = logoShape.Y; // pixels from worksheet top border

        // Log the coordinates for debugging
        Console.WriteLine($"Shape 'Logo' absolute coordinates: X = {absoluteX} px, Y = {absoluteY} px");

        // Save the workbook (lifecycle rule)
        workbook.Save("ShapeCoordinatesDemo.xlsx");
    }
}