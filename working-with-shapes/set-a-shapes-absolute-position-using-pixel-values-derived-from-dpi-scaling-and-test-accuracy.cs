// Title: Set a rectangle shape's absolute position in an Excel worksheet using pixel coordinates and DPI scaling with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape left and top using pixel values | convert screen pixels to points for Excel shape placement Aspose | how to position a rectangle shape at exact pixel coordinates in Aspose.Cells | apply DPI scaling when positioning shapes in an Excel file with Aspose.Cells | verify shape coordinates after setting Left and Top in Aspose.Cells C#
// Tags: pixel-to-point conversion Aspose.Cells C# | absolute shape positioning Excel Aspose | DPI scaling for Excel shapes | rectangle shape placement Aspose.Cells | shape coordinate verification Aspose.Cells | set shape Left Top properties C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositionExample
{
    // The example creates a new workbook, adds a rectangle shape to the first worksheet, converts desired pixel offsets (200 px X, 150 px Y) to points using a 96 DPI scaling factor, assigns the calculated points to the shape's Left and Top properties, checks the values against a small tolerance, and saves the workbook as ShapePositionTest.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    0, // upper left row
                    0, // upper left column
                    0, // upper left row offset (pixels)
                    0, // upper left column offset (pixels)
                    100, // height (points)
                    100); // width (points)

                // Desired absolute position in pixels
                int pixelX = 200; // horizontal pixel offset
                int pixelY = 150; // vertical pixel offset

                // DPI scaling factor (typical screen DPI is 96)
                double dpi = 96.0;

                // Convert pixel values to points (1 point = 1/72 inch)
                double pointsPerPixel = 72.0 / dpi;
                double leftInPoints = pixelX * pointsPerPixel;
                double topInPoints = pixelY * pointsPerPixel;

                // Set the shape's absolute position (cast to int if the API expects integer values)
                shape.Left = (int)Math.Round(leftInPoints);
                shape.Top = (int)Math.Round(topInPoints);

                // Verify that the position was set correctly
                double actualLeft = shape.Left;
                double actualTop = shape.Top;
                double tolerance = 0.001; // tolerance for floating‑point comparison

                bool leftMatches = Math.Abs(actualLeft - leftInPoints) < tolerance;
                bool topMatches = Math.Abs(actualTop - topInPoints) < tolerance;

                Console.WriteLine($"Expected Left (pts): {leftInPoints}, Actual Left (pts): {actualLeft}, Match: {leftMatches}");
                Console.WriteLine($"Expected Top (pts): {topInPoints}, Actual Top (pts): {actualTop}, Match: {topMatches}");

                // Save the workbook to verify the shape appears at the intended location
                string outputPath = "ShapePositionTest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
