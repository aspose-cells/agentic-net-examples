// Title: Aspose.Cells C# – Retrieve and Sort Worksheet Shapes by Z‑Order
// Description: Demonstrates how to create a workbook, add shapes, obtain the ShapeCollection, sort the shapes by their ZOrderPosition property, and print each shape's name, type, and Z‑order index to the console before saving the file.
// Keywords: Aspose.Cells | C# | shape ZOrderPosition | sort worksheet shapes | ShapeCollection | Excel shape ordering | Aspose.Cells example
// Common Searches: Aspose.Cells sort shapes by Z-order | C# get shape ZOrderPosition | list Excel shapes with Aspose.Cells | retrieve shape collection .NET | order shapes in worksheet Aspose
// Developer Intent: Get every shape from a worksheet, order them by Z‑order, and display the sorted information.
// Use Cases: Create a console report of shape layering to troubleshoot overlapping objects. | Reorder shapes programmatically before exporting or printing the workbook. | Export shape metadata (name, type, Z‑order) for documentation or integration with other systems.
// AI Prompts: Generate C# code using Aspose.Cells that lists all worksheet shapes sorted by ZOrderPosition. | Show how to log each shape's name, type, and Z‑order after sorting with Aspose.Cells. | Explain how to modify a shape's Z‑order after retrieving and sorting it in C#.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeZOrderDemo
{
    // Demonstrates how to create a workbook, add shapes, obtain the ShapeCollection, sort the shapes by their ZOrderPosition property, and print each shape's name, type, and Z‑order index to the console before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample shapes (for demonstration purposes)
            sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 100);
            sheet.Shapes.AddOval(5, 0, 5, 0, 100, 100);
            sheet.Shapes.AddLine(8, 0, 8, 0, 100, 100);

            // Retrieve the shape collection
            ShapeCollection shapes = sheet.Shapes;

            // Build a list of shapes and sort them by ZOrderPosition
            List<Shape> sortedShapes = shapes.Cast<Shape>()
                                             .OrderBy(s => s.ZOrderPosition)
                                             .ToList();

            // Output the sorted shapes to the console
            Console.WriteLine("Shapes sorted by Z-order position:");
            foreach (Shape shape in sortedShapes)
            {
                // Display shape name (if set) and its Z-order index
                string name = string.IsNullOrEmpty(shape.Name) ? "(no name)" : shape.Name;
                Console.WriteLine($"Name: {name}, Type: {shape.Type}, ZOrderPosition: {shape.ZOrderPosition}");
            }

            // Save the workbook (using the standard save rule)
            workbook.Save("SortedShapesDemo.xlsx");
        }
    }
}
