using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeAbsolutePositionWithDpi
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define DPI (dots per inch). Typical screen DPI is 96.
            const double dpi = 96.0;

            // Desired offset from the worksheet's top-left corner in centimeters
            double offsetCmX = 2.5; // 2.5 cm to the right
            double offsetCmY = 1.8; // 1.8 cm down

            // Convert centimeters to inches (1 inch = 2.54 cm)
            double offsetInchX = offsetCmX / 2.54;
            double offsetInchY = offsetCmY / 2.54;

            // Convert inches to pixels using the DPI value
            int offsetPixelX = (int)Math.Round(offsetInchX * dpi);
            int offsetPixelY = (int)Math.Round(offsetInchY * dpi);

            // Add a rectangle shape (row, column, top offset, left offset, height, width)
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 150);

            // Set the absolute position using pixel values derived from DPI scaling
            shape.X = offsetPixelX; // horizontal offset from worksheet left border (pixels)
            shape.Y = offsetPixelY; // vertical offset from worksheet top border (pixels)

            // Verify the position by reading back the properties
            Console.WriteLine($"Desired offset (cm): X={offsetCmX} cm, Y={offsetCmY} cm");
            Console.WriteLine($"Converted offset (pixels): X={offsetPixelX} px, Y={offsetPixelY} px");
            Console.WriteLine($"Shape.X (pixels): {shape.X}");
            Console.WriteLine($"Shape.Y (pixels): {shape.Y}");
            Console.WriteLine($"Shape.Left (pixels): {shape.Left}");
            Console.WriteLine($"Shape.Top (pixels): {shape.Top}");

            // Save the workbook
            string outputFile = "ShapeAbsolutePositionWithDpi.xlsx";
            try
            {
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}