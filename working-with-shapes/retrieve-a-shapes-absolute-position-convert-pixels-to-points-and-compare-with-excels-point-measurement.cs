// Title: Get Shape Absolute Position, Convert Pixels to Points, and Compare with Excel Points using Aspose.Cells for .NET
// Description: Shows how to add a rectangle, read its X/Y pixel offsets, translate them to points (assuming 96 dpi), obtain WidthPt/HeightPt, and validate pixel‑derived dimensions against Excel's point measurements in a C# workbook.
// Keywords: Aspose.Cells shape position | pixel to point conversion | WidthPt HeightPt | absolute shape coordinates | Excel points vs pixels | C# Aspose.Cells example | shape dimension scaling | DPI conversion
// Common Searches: aspnet get shape X Y offset Aspose.Cells | convert shape pixels to points C# | compare shape width in points and pixels Aspose | retrieve shape dimensions in points from workbook | pixel to point conversion formula Aspose.Cells
// Developer Intent: The developer needs the exact pixel location of a shape, a conversion to typographic points, and a way to confirm those values match Excel's native point properties.
// Use Cases: Align shapes precisely when exporting to PDF by using point units. | Ensure program‑generated graphics replicate hand‑crafted Excel layouts. | Adapt shape placement dynamically for displays with different DPI settings.
// AI Prompts: Write C# code with Aspose.Cells that reads a shape's X and Y pixel offsets and converts them to points for a 96 dpi screen. | Provide a snippet that compares a shape's WidthPt property with a width calculated from its pixel width. | Show how to shift a shape by a given point offset after converting its original pixel coordinates.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle, read its X/Y pixel offsets, translate them to points (assuming 96 dpi), obtain WidthPt/HeightPt, and validate pixel‑derived dimensions against Excel's point measurements in a C# workbook.
class RetrieveShapePosition
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape.
        // Parameters: upper left row, top offset (pixels), left column, left offset (pixels), height (pixels), width (pixels)
        Shape rectangle = worksheet.Shapes.AddRectangle(5, 100, 2, 50, 200, 150);

        // ----- Retrieve absolute position in pixels -----
        // X and Y give the horizontal and vertical offset of the shape from the worksheet's left/top border (in pixels)
        int xPixel = rectangle.X;
        int yPixel = rectangle.Y;

        // ----- Convert pixels to points -----
        // Excel uses points (1 point = 1/72 inch). Assuming a screen DPI of 96 (standard for Aspose.Cells),
        // points = pixels * 72 / DPI
        const double dpi = 96.0;
        double xPoint = xPixel * 72.0 / dpi;
        double yPoint = yPixel * 72.0 / dpi;

        // ----- Retrieve shape size in points (for comparison) -----
        double widthPt = rectangle.WidthPt;   // width in points
        double heightPt = rectangle.HeightPt; // height in points

        // Output the results
        Console.WriteLine($"Shape X offset: {xPixel} px = {xPoint:F2} pt");
        Console.WriteLine($"Shape Y offset: {yPixel} px = {yPoint:F2} pt");
        Console.WriteLine($"Shape width: {widthPt} pt, height: {heightPt} pt");

        // Example comparison: verify that the pixel‑derived width matches the point width
        // (rectangle.Width is in pixels; convert to points using the same factor)
        double widthFromPixelsPt = rectangle.Width * 72.0 / dpi;
        Console.WriteLine($"Width from pixels: {widthFromPixelsPt:F2} pt (should match WidthPt)");

        // Save the workbook
        workbook.Save("ShapePositionDemo.xlsx");
    }
}
