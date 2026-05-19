using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample shapes to the worksheet
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
        sheet.Shapes.AddRectangle(2, 10, 2, 10, 100, 200); // Y = 10
        sheet.Shapes.AddOval(5, 30, 5, 30, 80, 80);       // Y = 30
        sheet.Shapes.AddLine(8, 20, 8, 20, 150, 150);    // Y = 20

        // Retrieve all shapes from the worksheet (Shapes property rule)
        ShapeCollection shapes = sheet.Shapes;

        // Collect each shape with its Y coordinate
        List<(Shape shape, int y)> shapeList = new List<(Shape, int)>();
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape s = shapes[i];               // indexer rule
            shapeList.Add((s, s.Y));           // Y property rule
        }

        // Sort shapes by Y coordinate (top to bottom)
        var sorted = shapeList.OrderBy(item => item.y).ToList();

        // Output the sorted order
        Console.WriteLine("Shapes sorted by Y coordinate (top to bottom):");
        for (int i = 0; i < sorted.Count; i++)
        {
            Shape s = sorted[i].shape;
            Console.WriteLine($"Order {i}: Name=\"{s.Name}\", Y={sorted[i].y}");
        }

        // Save the workbook (save rule)
        workbook.Save("ShapesSortedByY.xlsx");
    }
}