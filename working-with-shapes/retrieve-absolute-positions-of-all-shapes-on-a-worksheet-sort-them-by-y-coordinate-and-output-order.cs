// Title: Retrieve absolute positions of worksheet shapes, sort them by Y (Top) coordinate, and output the sorted order with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console program that loads an Excel workbook using Aspose.Cells, iterates through the ShapeCollection of the first worksheet, records each shape's Top and Left values, sorts the shapes by the Top coordinate, and prints the shape index with its coordinates. | Generate C# code that verifies the existence of an input .xlsx file, extracts absolute positions of all shapes on a worksheet via Aspose.Cells, orders the shapes vertically, and includes robust exception handling. | Create a reusable C# method that accepts a Worksheet object, returns a list of shape indexes sorted by their Y position, and demonstrate its usage with Aspose.Cells.
// Common Searches: C# Aspose.Cells get shape Top and Left coordinates from worksheet | How to sort Excel shapes by vertical position using Aspose.Cells .NET | List all shapes on a worksheet with absolute positions in Aspose.Cells | Retrieve and order shape indexes by Y coordinate in an Excel file with C# | Aspose.Cells shape collection enumeration and sorting example
// Tags: Aspose.Cells retrieve shape coordinates | C# sort worksheet shapes by Y position | enumerate ShapeCollection absolute positions | Aspose.Cells shape sorting example | console output shape index with coordinates

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads input.xlsx, reads all shapes from the first worksheet, captures each shape's Top (Y) and Left (X) coordinates, sorts the shapes ascending by Top, and prints the sorted shape index along with its coordinates.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve all shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Store shape index and its absolute position (Top = Y, Left = X)
            var shapeInfos = new List<(int Index, double Top, double Left)>();

            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];

                // Top and Left provide absolute coordinates in points
                shapeInfos.Add((i, shape.Top, shape.Left));
            }

            // Sort shapes by Y coordinate (Top) ascending
            var sortedShapes = shapeInfos.OrderBy(s => s.Top).ToList();

            // Output the sorted order
            Console.WriteLine("Shapes sorted by Y coordinate (Top):");
            foreach (var info in sortedShapes)
            {
                Console.WriteLine($"Shape Index: {info.Index}, Top: {info.Top}, Left: {info.Left}");
            }

            // Uncomment and modify the path if you need to save changes
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
