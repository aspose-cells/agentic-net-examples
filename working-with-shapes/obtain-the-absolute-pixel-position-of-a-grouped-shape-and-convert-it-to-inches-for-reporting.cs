// Title: Get GroupShape absolute pixel position and convert to inches using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a rectangle and an oval, group them, read the GroupShape X and Y pixel offsets, convert those values to inches (96 dpi), optionally use the LeftInch/TopInch shortcuts, display both units, and save the workbook.
// Keywords: Aspose.Cells | GroupShape | pixel position | inch conversion | C# .NET | shape coordinates | absolute position | 96 DPI | LeftInch | TopInch | shape layout reporting
// Common Searches: Aspose.Cells get GroupShape pixel coordinates | Convert GroupShape position from pixels to inches C# | GroupShape X Y properties Aspose | How to read absolute shape location in Aspose.Cells | Pixel to inch conversion for shapes Aspose.Cells
// Developer Intent: Obtain the absolute pixel coordinates of a grouped shape and express them in inches.
// Use Cases: Create a layout audit that lists each grouped shape’s position in pixels and inches for precise document verification. | Programmatically align grouped shapes across worksheets by comparing their inch measurements and adjusting offsets. | Generate a PDF export that preserves exact shape placement by using converted inch values for scaling.
// AI Prompts: Write C# code with Aspose.Cells that reads a GroupShape’s X and Y pixel values, converts them to centimeters, and logs the results. | Provide an example that extracts a GroupShape’s pixel coordinates, converts them to inches, writes the data to a summary worksheet, and saves the file. | Explain Aspose.Cells’ pixel‑to‑inch conversion logic for shapes and demonstrate when to use the LeftInch/TopInch properties instead of manual calculations.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupedShapePosition
{
    // Shows how to create a workbook, add a rectangle and an oval, group them, read the GroupShape X and Y pixel offsets, convert those values to inches (96 dpi), optionally use the LeftInch/TopInch shortcuts, display both units, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two sample shapes
            Shape rect1 = sheet.Shapes.AddRectangle(2, 1, 0, 100, 80, 0);
            Shape rect2 = sheet.Shapes.AddOval(4, 3, 0, 120, 90, 0);

            // Group the shapes
            GroupShape group = sheet.Shapes.Group(new Shape[] { rect1, rect2 });

            // Absolute pixel position of the group shape (X and Y are in pixels)
            double groupPosXPixel = group.X; // horizontal offset from worksheet left border
            double groupPosYPixel = group.Y; // vertical offset from worksheet top border

            // Convert pixels to inches (Aspose uses 96 DPI for conversion)
            double groupPosXInch = groupPosXPixel / 96.0;
            double groupPosYInch = groupPosYPixel / 96.0;

            // Alternatively, you can directly use the provided inch properties
            // double groupPosXInchAlt = group.LeftInch;
            // double groupPosYInchAlt = group.TopInch;

            // Output the results
            Console.WriteLine($"Group Shape Position (pixels): X = {groupPosXPixel}, Y = {groupPosYPixel}");
            Console.WriteLine($"Group Shape Position (inches): X = {groupPosXInch:F2}, Y = {groupPosYInch:F2}");

            // Save the workbook (save rule)
            workbook.Save("GroupedShapePositionReport.xlsx");
        }
    }
}
