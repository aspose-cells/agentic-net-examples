// Title: Position a rectangle shape at specific pixel coordinates in an Excel worksheet using Aspose.Cells for .NET with DPI conversion
// AI Prompts: Compute point values from a 200 px horizontal offset and a 150 px vertical offset using a 96 DPI factor, then assign those values to the Shape.Left and Shape.Top properties to place the rectangle precisely. | Read the Shape.Left and Shape.Top after placement, convert them back to pixels, calculate the deviation from the original pixel offsets, and report whether the difference is within a 0.5‑pixel tolerance.
// Common Searches: Aspose.Cells set shape position using pixel coordinates | convert pixel offset to points for shape placement in Excel .NET | how to verify shape location accuracy after setting Left and Top in Aspose.Cells | apply DPI scaling when positioning shapes in an Excel workbook with Aspose.Cells | C# Aspose.Cells shape absolute positioning tolerance test
// Tags: pixel-to-point conversion Aspose.Cells | shape left top properties .NET | absolute shape positioning Excel | DPI scaling shape coordinates | shape placement tolerance verification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositionDemo
{
    // The example creates a workbook, adds a rectangle shape, converts a 200 px × 150 px offset to points using a 96 DPI factor, sets the shape's Left and Top properties, converts the stored values back to pixels, compares them with the requested coordinates within a 0.5‑pixel tolerance, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define DPI (dots per inch) used for pixel‑to‑point conversion.
                const double dpi = 96.0;

                // Desired absolute position in pixels
                double pixelX = 200; // horizontal offset from the left edge
                double pixelY = 150; // vertical offset from the top edge

                // Convert pixel values to points (1 point = 1/72 inch)
                double pointsX = pixelX * 72.0 / dpi;
                double pointsY = pixelY * 72.0 / dpi;

                // Add a rectangle shape to the worksheet.
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width (all in points)
                Shape rect = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // correct enum for shape type
                    0, 0,                     // upper left cell (row, column)
                    0, 0,                     // top and left offsets (initially 0)
                    100, 100);                // height and width in points

                // Set the absolute position using the calculated point values (cast to int if required by the API)
                rect.Left = (int)pointsX;
                rect.Top = (int)pointsY;

                // ---- Test Accuracy ----
                // Retrieve the shape's stored position (in points) and convert back to pixels
                double actualPixelX = rect.Left * dpi / 72.0;
                double actualPixelY = rect.Top * dpi / 72.0;

                // Compute differences
                double diffX = Math.Abs(actualPixelX - pixelX);
                double diffY = Math.Abs(actualPixelY - pixelY);

                // Define an acceptable tolerance (e.g., 0.5 pixel)
                const double tolerance = 0.5;

                // Output results
                Console.WriteLine($"Requested position: ({pixelX} px, {pixelY} px)");
                Console.WriteLine($"Actual position   : ({actualPixelX:F2} px, {actualPixelY:F2} px)");
                Console.WriteLine($"Difference        : ({diffX:F2} px, {diffY:F2} px)");
                Console.WriteLine($"Within tolerance? : {(diffX <= tolerance && diffY <= tolerance ? "Yes" : "No")}");

                // Save the workbook to verify the shape visually (optional)
                string outputPath = "ShapePositionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
