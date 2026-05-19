using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

class ShadowEffectDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

        // Access the shape's shadow effect and configure it
        ShadowEffect shadow = shape.ShadowEffect;
        shadow.Angle = 45;          // Set angle to 45 degrees
        shadow.Distance = 10;       // Set distance to 10 points

        // Create a dark gray color and assign it to the shadow
        CellsColor darkGray = workbook.CreateCellsColor();
        darkGray.Color = Color.DarkGray;
        shadow.Color = darkGray;

        // Save the workbook with the configured shadow effect
        workbook.Save("ShadowEffectDemo.xlsx");
    }
}