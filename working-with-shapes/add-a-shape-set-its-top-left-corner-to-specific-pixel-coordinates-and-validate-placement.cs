using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePlacementDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired pixel coordinates for the shape's top-left corner
            int desiredTopPixel = 150;   // vertical offset in pixels
            int desiredLeftPixel = 200;  // horizontal offset in pixels

            // Add a rectangle shape at row 0, column 0 with the desired offsets
            // Height and width are arbitrary (e.g., 100x80 pixels)
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,                        // topRow index
                desiredTopPixel,          // top offset (pixels)
                0,                        // leftColumn index
                desiredLeftPixel,         // left offset (pixels)
                80,                       // height (pixels)
                100);                     // width (pixels)

            // Validate that the shape's Top and Left properties match the expected values
            bool topMatches = shape.Top == desiredTopPixel;
            bool leftMatches = shape.Left == desiredLeftPixel;

            Console.WriteLine($"Top placement validation: {(topMatches ? "Passed" : "Failed")}");
            Console.WriteLine($"Left placement validation: {(leftMatches ? "Passed" : "Failed")}");

            // Save the workbook to verify the shape appears at the specified location
            workbook.Save("ShapePlacementResult.xlsx");
        }
    }
}