using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

            // Set the shape's Z-order position to zero (default layer)
            shape.ZOrderPosition = 0; // using Shape.ZOrderPosition property rule

            // Verify that the Z-order position is set to zero
            if (shape.ZOrderPosition == 0)
            {
                Console.WriteLine("Shape Z-order is correctly set to the default position (0).");
            }
            else
            {
                Console.WriteLine($"Unexpected Z-order position: {shape.ZOrderPosition}");
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save("ShapeZOrderDemo.xlsx");
        }
    }
}