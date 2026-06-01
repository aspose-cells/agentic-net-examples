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
            Worksheet worksheet = workbook.Worksheets[0];

            // Add three overlapping rectangle shapes
            Shape shape1 = worksheet.Shapes.AddRectangle(5, 5, 5, 5, 120, 120);   // Bottom‑most
            Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 30, 30, 120, 120);
            Shape shape3 = worksheet.Shapes.AddRectangle(55, 55, 55, 55, 120, 120); // Top‑most initially

            // Assign explicit Z‑order positions
            shape1.ZOrderPosition = 0; // Back
            shape2.ZOrderPosition = 1;
            shape3.ZOrderPosition = 2; // Front

            // Output current Z‑order positions
            Console.WriteLine($"Initial ZOrderPosition: shape1={shape1.ZOrderPosition}, shape2={shape2.ZOrderPosition}, shape3={shape3.ZOrderPosition}");

            // Bring shape1 to the front using ToFrontOrBack (positive value moves forward)
            shape1.ToFrontOrBack(3); // Move it ahead of the other two shapes

            // After moving, output the updated Z‑order positions
            Console.WriteLine($"After ToFrontOrBack: shape1={shape1.ZOrderPosition}, shape2={shape2.ZOrderPosition}, shape3={shape3.ZOrderPosition}");

            // Save the workbook to visualize the rendering order
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}