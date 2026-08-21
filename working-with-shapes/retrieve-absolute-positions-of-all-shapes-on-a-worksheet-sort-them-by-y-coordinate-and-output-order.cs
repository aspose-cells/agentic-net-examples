// Title: Aspose.Cells .NET – Retrieve and Sort Worksheet Shapes by Absolute Y Position
// Description: Demonstrates how to enumerate all shapes on an Aspose.Cells worksheet, read each shape's absolute Y coordinate (pixel offset from the top), sort the shapes from top to bottom, display their names and Y values, and optionally save the workbook.
// Keywords: Aspose.Cells | C# | shape Y coordinate | absolute shape position | sort worksheet shapes | Excel shape ordering | retrieve shape locations | Aspose.Cells Drawing | Excel automation
// Common Searches: Aspose.Cells get shape Y coordinate .NET | sort Excel shapes by vertical position using Aspose | list all shapes with absolute positions in a workbook | C# retrieve shape locations Aspose.Cells | order shapes top to bottom Aspose.Cells
// Developer Intent: Extract the absolute Y positions of every shape on a worksheet, arrange them in ascending Y order, and output the sorted sequence.
// Use Cases: Create a top‑to‑bottom inventory of chart, image, and drawing objects for reporting or auditing. | Adjust drawing order before exporting to PDF or image to ensure correct visual layering. | Programmatically align or reposition shapes based on their vertical layout within a sheet.
// AI Prompts: Generate C# code with Aspose.Cells that collects each shape's absolute Y coordinate, sorts the shapes from top to bottom, and prints their names and Y values. | Show how to output the sorted shape list and then save the workbook as an Excel file using Aspose.Cells. | Explain how to extend the example to also sort shapes by X coordinate after the Y‑coordinate sort.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to enumerate all shapes on an Aspose.Cells worksheet, read each shape's absolute Y coordinate (pixel offset from the top), sort the shapes from top to bottom, display their names and Y values, and optionally save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Add sample shapes (for demonstration purposes)
        // -------------------------------------------------
        // Rectangle shape
        worksheet.Shapes.AddRectangle(2, 10, 2, 10, 50, 100);
        // Oval shape
        worksheet.Shapes.AddOval(5, 30, 5, 30, 60, 80);
        // Line shape
        worksheet.Shapes.AddLine(8, 20, 8, 20, 100, 5);

        // -------------------------------------------------
        // Retrieve all shapes with their absolute Y positions
        // -------------------------------------------------
        var shapeInfos = new List<(Shape shape, int y)>();
        for (int i = 0; i < worksheet.Shapes.Count; i++)
        {
            Shape shape = worksheet.Shapes[i];
            // Y property gives the vertical offset from the worksheet top border (pixels)
            shapeInfos.Add((shape, shape.Y));
        }

        // -------------------------------------------------
        // Sort shapes by Y coordinate (top to bottom)
        // -------------------------------------------------
        var sortedShapes = shapeInfos.OrderBy(info => info.y).ToList();

        // -------------------------------------------------
        // Output the sorted order
        // -------------------------------------------------
        Console.WriteLine("Shapes sorted by Y coordinate (top to bottom):");
        for (int i = 0; i < sortedShapes.Count; i++)
        {
            Shape shape = sortedShapes[i].shape;
            int yPos = sortedShapes[i].y;
            Console.WriteLine($"Order {i + 1}: Name = \"{shape.Name}\", Y = {yPos}");
        }

        // -------------------------------------------------
        // Save the workbook (optional)
        // -------------------------------------------------
        workbook.Save("ShapesSortedByY.xlsx");
    }
}
