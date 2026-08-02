// Title: Get GroupShape X/Y in pixels and convert to inches with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a rectangle and an oval, group them, read the GroupShape's X and Y pixel offsets from the worksheet origin, convert the offsets to inches using the default 96 DPI, and display the results.
// Keywords: Aspose.Cells GroupShape position | shape pixel to inch conversion | C# Aspose.Cells shape coordinates | default DPI 96 Aspose.Cells | retrieve X Y of grouped shape
// Common Searches: Aspose.Cells get GroupShape pixel location | convert shape coordinates to inches .NET | how to read X Y of a grouped shape in Excel | pixel to inch conversion Aspose.Cells | default DPI used by Aspose.Cells for shapes
// Developer Intent: Obtain the absolute pixel offsets of a GroupShape and express them in inches.
// Use Cases: Document exact placement of grouped graphics for print‑ready reports. | Compare layout positions across worksheets when aligning multiple groups. | Generate a layout audit that lists shape locations in both pixels and inches.
// AI Prompts: Write C# code with Aspose.Cells that reads a GroupShape's X and Y values and converts them to centimeters. | Explain the relationship between pixel coordinates, DPI, and physical units in Aspose.Cells shape measurements. | Provide an example that groups several shapes, logs their absolute positions in pixels, inches, and millimeters.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a rectangle and an oval, group them, read the GroupShape's X and Y pixel offsets from the worksheet origin, convert the offsets to inches using the default 96 DPI, and display the results.
class GroupShapePositionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two sample shapes that will be grouped
        Shape rect1 = sheet.Shapes.AddRectangle(2, 2, 0, 100, 80, 0);
        Shape rect2 = sheet.Shapes.AddOval(5, 5, 0, 120, 90, 0);

        // Group the shapes
        GroupShape group = sheet.Shapes.Group(new Shape[] { rect1, rect2 });

        // Absolute position of the group shape in pixels (relative to worksheet top‑left corner)
        int absoluteXPixel = (int)group.X; // horizontal offset in pixels
        int absoluteYPixel = (int)group.Y; // vertical offset in pixels

        // Convert pixels to inches (Aspose.Cells uses 96 DPI by default)
        double dpi = 96.0;
        double absoluteXInch = absoluteXPixel / dpi;
        double absoluteYInch = absoluteYPixel / dpi;

        // Output the results
        Console.WriteLine($"Group Shape Position (pixels): X = {absoluteXPixel}, Y = {absoluteYPixel}");
        Console.WriteLine($"Group Shape Position (inches): X = {absoluteXInch:F2}, Y = {absoluteYInch:F2}");

        // Save the workbook (optional, just to demonstrate lifecycle compliance)
        workbook.Save("GroupShapePositionDemo.xlsx");
    }
}
