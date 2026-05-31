using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapePositionComparison
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape with initial size (width:100, height:100) at row 1, column 0
                // Parameters: topRow, top, leftColumn, left, height, width, rotationAngle
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 0);

                // Set shape position in pixels (horizontal offset from worksheet left border, vertical offset from top row)
                shape.X = 150; // pixels
                shape.Y = 200; // pixels

                // Set shape size using points (1 point = 1/72 inch)
                shape.WidthPt = 120; // points
                shape.HeightPt = 80; // points

                // Retrieve the absolute position in pixels
                int xPixels = shape.X;
                int yPixels = shape.Y;

                // Convert pixels to points.
                // Excel assumes 96 DPI, so 1 pixel = 72/96 points = 0.75 points.
                double xPointsFromPixels = xPixels * 72.0 / 96.0;
                double yPointsFromPixels = yPixels * 72.0 / 96.0;

                // Output the values
                Console.WriteLine($"Shape X position: {xPixels} pixels = {xPointsFromPixels:F2} points");
                Console.WriteLine($"Shape Y position: {yPixels} pixels = {yPointsFromPixels:F2} points");
                Console.WriteLine($"Shape Width: {shape.WidthPt} points");
                Console.WriteLine($"Shape Height: {shape.HeightPt} points");

                // Compare pixel‑derived points with the shape's point measurements
                if (Math.Abs(xPointsFromPixels - shape.WidthPt) < 0.01)
                    Console.WriteLine("Horizontal pixel‑to‑point conversion matches shape width.");
                else
                    Console.WriteLine("Horizontal conversion does NOT match shape width.");

                if (Math.Abs(yPointsFromPixels - shape.HeightPt) < 0.01)
                    Console.WriteLine("Vertical pixel‑to‑point conversion matches shape height.");
                else
                    Console.WriteLine("Vertical conversion does NOT match shape height.");

                // Save the workbook (ensure the directory exists)
                string outputPath = "ShapePositionComparison.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}