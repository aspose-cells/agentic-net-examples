using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyOuterShadow
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Get the shadow effect object of the shape
        ShadowEffect shadow = shape.ShadowEffect;

        // Apply an outer preset shadow (e.g., OffsetBottom)
        shadow.PresetType = PresetShadowType.OffsetBottom;

        // Set the transparency of the shadow to 40% (0.4)
        shadow.Transparency = 0.4;

        // Save the workbook to a file
        workbook.Save("ShapeWithOuterShadow.xlsx");
    }
}