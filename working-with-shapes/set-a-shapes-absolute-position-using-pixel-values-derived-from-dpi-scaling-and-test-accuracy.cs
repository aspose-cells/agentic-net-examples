// Title: Set a shape's absolute position with DPI‑scaled pixels in Aspose.Cells (C#)
// Description: Creates a workbook, adds a rectangle shape, converts target inches to pixel values using a 96 DPI factor, assigns the pixel values to the shape's X and Y properties, cross‑checks with LeftInch/TopInch, prints expected vs. actual coordinates, and saves the file.
// Keywords: Aspose.Cells shape position pixels | DPI scaling Aspose.Cells | shape X Y properties C# | convert inches to pixels Aspose.Cells | verify shape coordinates | absolute shape placement | Aspose.Cells rectangle example
// Common Searches: Aspose.Cells set shape position in pixels | how to use DPI to place shapes in Aspose.Cells | convert inches to pixels for shape placement C# | verify shape X Y values Aspose.Cells | absolute shape coordinates Aspose.Cells .NET
// Developer Intent: Place a shape at exact worksheet coordinates by calculating pixel values from inches using DPI, then confirm the placement through both pixel and inch properties.
// Use Cases: Position a company logo 2.5 inches from the left and 1.75 inches from the top for printable reports. | Align multiple graphics consistently across sheets by converting design measurements to pixels based on workbook DPI. | Validate that programmatic shape placement matches design specifications by comparing X/Y and LeftInch/TopInch values.
// AI Prompts: Generate C# code that positions a circle shape 3 inches from the left and 2 inches from the top using DPI‑based pixel conversion in Aspose.Cells. | Write a method that receives inch coordinates and returns pixel values for shape placement, handling custom DPI settings. | Explain how to read back and compare a shape's pixel (X, Y) and inch (LeftInch, TopInch) positions after setting them in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle shape, converts target inches to pixel values using a 96 DPI factor, assigns the pixel values to the shape's X and Y properties, cross‑checks with LeftInch/TopInch, prints expected vs. actual coordinates, and saves the file.
    public class ShapeAbsolutePositionDpiDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (initial position and size are placeholders)
                // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
                Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 100);

                // Define DPI (dots per inch). Aspose.Cells uses 96 DPI by default for screen rendering.
                const double dpi = 96.0;

                // Desired absolute position in inches (example values)
                double targetInchesX = 2.5; // 2.5 inches from the left border
                double targetInchesY = 1.75; // 1.75 inches from the top border

                // Convert inches to pixel values using DPI scaling
                int targetPixelsX = (int)Math.Round(targetInchesX * dpi);
                int targetPixelsY = (int)Math.Round(targetInchesY * dpi);

                // Set the shape's absolute position using pixel properties (X and Y)
                shape.X = targetPixelsX;
                shape.Y = targetPixelsY;

                // Verify the position by reading back the pixel values
                int actualPixelsX = shape.X;
                int actualPixelsY = shape.Y;

                // Output the expected and actual pixel positions
                Console.WriteLine($"Expected X (pixels): {targetPixelsX}, Actual X (pixels): {actualPixelsX}");
                Console.WriteLine($"Expected Y (pixels): {targetPixelsY}, Actual Y (pixels): {actualPixelsY}");

                // Additionally, demonstrate setting position via inches properties for cross‑check
                shape.LeftInch = targetInchesX;
                shape.TopInch = targetInchesY;

                // Read back inches to ensure consistency
                double actualInchesX = shape.LeftInch;
                double actualInchesY = shape.TopInch;
                Console.WriteLine($"Expected X (inches): {targetInchesX}, Actual X (inches): {actualInchesX}");
                Console.WriteLine($"Expected Y (inches): {targetInchesY}, Actual Y (inches): {actualInchesY}");

                // Save the workbook to verify that the shape is positioned as intended
                workbook.Save("ShapeAbsolutePositionDpiDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeAbsolutePositionDpiDemo.Run();
        }
    }
}
