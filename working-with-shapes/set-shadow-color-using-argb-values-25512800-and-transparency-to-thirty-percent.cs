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

        // Add a rectangle shape to demonstrate the shadow effect
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 100);

        // Retrieve the shadow effect of the shape
        ShadowEffect shadow = shape.ShadowEffect;

        // Create a CellsColor instance and set its ARGB value to (255, 128, 0, 0)
        CellsColor shadowColor = workbook.CreateCellsColor();
        shadowColor.Argb = Color.FromArgb(255, 128, 0, 0).ToArgb(); // Opaque dark red

        // Apply the color to the shadow effect
        shadow.Color = shadowColor;

        // Set the shadow transparency to 30% (0.3)
        shadow.Transparency = 0.3;

        // Save the workbook
        workbook.Save("ShadowEffectARGB.xlsx");
    }
}