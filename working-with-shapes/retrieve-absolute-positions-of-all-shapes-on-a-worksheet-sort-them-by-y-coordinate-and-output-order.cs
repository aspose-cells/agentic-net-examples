// Title: C# – Get absolute positions of all worksheet shapes and sort by Y coordinate with Aspose.Cells
// Description: Shows how to enumerate every shape on an Aspose.Cells worksheet, read each shape's absolute X and Y coordinates, sort the collection from top‑to‑bottom, output the sorted order, and save the workbook.
// Keywords: Aspose.Cells shape Y coordinate | C# sort worksheet shapes | retrieve shape positions Aspose | Excel shape absolute location .NET | shape collection ordering | Aspose.Cells GetTopLeftRow
// Common Searches: Aspose.Cells get shape Y position | How to sort Excel shapes by vertical position in C# | List all shapes on a worksheet using Aspose.Cells | Retrieve shape coordinates without Excel Interop | C# order shapes top to bottom Aspose.Cells
// Developer Intent: Obtain the absolute X/Y positions of every shape on a worksheet, sort the shapes by their Y coordinate, and display the sorted order.
// Use Cases: Create a layout index that lists diagram elements from top to bottom for visual analysis. | Programmatically re‑position shapes after sorting to achieve vertical alignment. | Export the sorted shape list (name and coordinates) to another system for further processing.
// AI Prompts: Write C# code using Aspose.Cells that returns a List<string> of shape names ordered from top to bottom. | Show how to access Shape.X and Shape.Y properties and sort the collection without using LINQ. | Generate a method that moves each shape to a new row after sorting them by Y coordinate.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to enumerate every shape on an Aspose.Cells worksheet, read each shape's absolute X and Y coordinates, sort the collection from top‑to‑bottom, output the sorted order, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample shapes to the worksheet
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
        worksheet.Shapes.AddRectangle(2, 10, 2, 20, 100, 100); // Shape 0
        worksheet.Shapes.AddOval(5, 50, 5, 30, 80, 80);       // Shape 1
        worksheet.Shapes.AddLine(8, 30, 8, 40, 200, 5);       // Shape 2

        // Retrieve all shapes from the worksheet
        ShapeCollection shapeCollection = worksheet.Shapes;
        List<Shape> shapes = new List<Shape>();
        for (int i = 0; i < shapeCollection.Count; i++)
        {
            shapes.Add(shapeCollection[i]);
        }

        // Sort shapes by their Y coordinate (top to bottom)
        List<Shape> sortedShapes = shapes.OrderBy(s => s.Y).ToList();

        // Output the sorted order
        Console.WriteLine("Shapes sorted by Y coordinate (top to bottom):");
        for (int i = 0; i < sortedShapes.Count; i++)
        {
            Shape s = sortedShapes[i];
            Console.WriteLine($"Sorted Index: {i}, Name: {s.Name}, Y: {s.Y}, X: {s.X}");
        }

        // Save the workbook
        workbook.Save("SortedShapesDemo.xlsx");
    }
}
