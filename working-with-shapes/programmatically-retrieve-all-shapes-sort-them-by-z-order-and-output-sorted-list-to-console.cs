using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeZOrderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample shapes (optional, for demonstration)
            Shape shape1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 100);
            shape1.Name = "Rectangle 1";
            shape1.ZOrderPosition = 2;

            Shape shape2 = worksheet.Shapes.AddOval(5, 0, 5, 0, 100, 100);
            shape2.Name = "Oval 1";
            shape2.ZOrderPosition = 0;

            Shape shape3 = worksheet.Shapes.AddLine(8, 0, 8, 0, 100, 100);
            shape3.Name = "Line 1";
            shape3.ZOrderPosition = 1;

            // Retrieve all shapes from the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Sort shapes by their Z-order position
            var sortedShapes = shapes.Cast<Shape>()
                                     .OrderBy(s => s.ZOrderPosition)
                                     .ToList();

            // Output the sorted list to the console
            Console.WriteLine("Shapes sorted by Z-order position:");
            foreach (Shape s in sortedShapes)
            {
                Console.WriteLine($"Name: {s.Name}, Type: {s.Type}, ZOrderPosition: {s.ZOrderPosition}");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("SortedShapesDemo.xlsx");
        }
    }
}