using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add three overlapping rectangle shapes
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape1 = sheet.Shapes.AddRectangle(10, 10, 0, 0, 100, 100);
            Shape shape2 = sheet.Shapes.AddRectangle(30, 30, 0, 0, 100, 100);
            Shape shape3 = sheet.Shapes.AddRectangle(50, 50, 0, 0, 100, 100);

            // Set initial Z-order positions (0 = backmost)
            shape1.ZOrderPosition = 0; // back
            shape2.ZOrderPosition = 1;
            shape3.ZOrderPosition = 2; // front

            Console.WriteLine($"Initial ZOrder: shape1={shape1.ZOrderPosition}, shape2={shape2.ZOrderPosition}, shape3={shape3.ZOrderPosition}");

            // Incrementally bring shape1 forward in the Z-order
            shape1.ToFrontOrBack(1); // move one position forward
            Console.WriteLine($"After moving shape1 forward 1 step: shape1 ZOrder={shape1.ZOrderPosition}");

            shape1.ToFrontOrBack(1); // move another position forward
            Console.WriteLine($"After moving shape1 forward another step: shape1 ZOrder={shape1.ZOrderPosition}");

            // Send shape3 backward in the Z-order
            shape3.ToFrontOrBack(-2); // move two positions back
            Console.WriteLine($"After sending shape3 back 2 steps: shape3 ZOrder={shape3.ZOrderPosition}");

            // Save the workbook to observe the layering effect
            workbook.Save("ZOrderDemo.xlsx");
        }
    }
}