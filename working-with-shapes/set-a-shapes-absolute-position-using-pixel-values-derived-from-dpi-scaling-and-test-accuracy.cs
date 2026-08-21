// Title: Position a Shape Precisely Using DPI‑Based Pixels in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle shape, converts inch offsets to pixel values using a 96 DPI scale, assigns the pixel values to the shape's X and Y properties, validates the placement, and saves the file.
// Keywords: Aspose.Cells shape position pixels | C# shape X Y Aspose.Cells | convert inches to pixels Aspose.Cells | absolute shape coordinates .NET | DPI scaling Excel shape placement
// Common Searches: Aspose.Cells set shape location in pixels | convert inches to pixels for Excel shapes C# | verify shape X and Y values Aspose.Cells | absolute positioning of rectangle shape Aspose.Cells
// Developer Intent: Set a shape's X/Y coordinates in pixel units derived from DPI conversion and confirm the placement.
// Use Cases: Place a logo exactly 2.5 inches from the left edge and 1.75 inches from the top edge of a worksheet. | Align multiple shapes consistently across workbooks by using DPI‑based pixel positioning. | Automate a test that asserts shape.X and shape.Y match expected pixel values after programmatic positioning.
// AI Prompts: Generate C# code that positions a shape at specific inch offsets using DPI conversion with Aspose.Cells. | Write a unit test that checks shape.X and shape.Y equal the calculated pixel values after setting them. | Explain how Aspose.Cells shape.X and shape.Y map to pixel coordinates and how DPI influences the conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePositionDemo
{
    // Creates a workbook, adds a rectangle shape, converts inch offsets to pixel values using a 96 DPI scale, assigns the pixel values to the shape's X and Y properties, validates the placement, and saves the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape with temporary position and size
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape shape = worksheet.Shapes.AddRectangle(
                    0,    // upperLeftRow
                    0,    // upperLeftColumn
                    0,    // top offset (pixels)
                    0,    // left offset (pixels)
                    100,  // height (pixels)
                    200   // width (pixels)
                );

                // Define DPI (dots per inch) for conversion. Typical screen DPI is 96.
                const double dpi = 96.0;

                // Desired absolute position in inches
                double desiredLeftInches = 2.5;   // 2.5 inches from the left border
                double desiredTopInches = 1.75;   // 1.75 inches from the top border

                // Convert inches to pixels using DPI scaling
                int expectedLeftPixels = (int)Math.Round(desiredLeftInches * dpi);
                int expectedTopPixels = (int)Math.Round(desiredTopInches * dpi);

                // Set the shape's absolute position using pixel properties X and Y
                shape.X = expectedLeftPixels;
                shape.Y = expectedTopPixels;

                // Verify that the shape's position matches the expected pixel values
                Console.WriteLine($"Expected X (pixels): {expectedLeftPixels}, Actual X: {shape.X}");
                Console.WriteLine($"Expected Y (pixels): {expectedTopPixels}, Actual Y: {shape.Y}");

                bool positionAccurate = shape.X == expectedLeftPixels && shape.Y == expectedTopPixels;
                Console.WriteLine($"Position accuracy test passed: {positionAccurate}");

                // Save the workbook (lifecycle save)
                string outputPath = "ShapeAbsolutePositionDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
