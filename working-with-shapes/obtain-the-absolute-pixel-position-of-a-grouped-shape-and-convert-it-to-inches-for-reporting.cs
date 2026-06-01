using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two sample shapes to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Shape rect = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 80);
        Shape oval = sheet.Shapes.AddOval(5, 5, 0, 0, 120, 90);

        // Group the shapes together
        GroupShape group = sheet.Shapes.Group(new Shape[] { rect, oval });

        // Obtain the absolute pixel position of the grouped shape
        int pixelX = group.X; // pixels from the worksheet's left border
        int pixelY = group.Y; // pixels from the worksheet's top border

        // Convert pixels to inches (Aspose.Cells assumes 96 DPI)
        double inchX = pixelX / 96.0;
        double inchY = pixelY / 96.0;

        // Alternatively, use the built‑in inch properties (they perform the same conversion)
        double leftInchProp = group.LeftInch;
        double topInchProp = group.TopInch;

        // Output the results
        Console.WriteLine($"Group position in pixels: X = {pixelX}, Y = {pixelY}");
        Console.WriteLine($"Group position in inches (calculated): Left = {inchX:F2}, Top = {inchY:F2}");
        Console.WriteLine($"Group position in inches (property): Left = {leftInchProp:F2}, Top = {topInchProp:F2}");

        // Save the workbook (required lifecycle step)
        workbook.Save("GroupedShapePosition.xlsx");
    }
}