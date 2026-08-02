using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

        // Change the glow size to 12 points
        shape.Glow.Size = 12;

        // Update the glow color to orange
        CellsColor glowColor = shape.Glow.Color;
        glowColor.Color = Color.Orange;

        // Save the workbook with the updated shape
        workbook.Save("ShapeGlowUpdated.xlsx");
    }
}