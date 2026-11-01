// Title: Retrieve absolute positions of all shapes on an Excel worksheet and sort them by Y coordinate using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through every Shape in a worksheet, captures each shape's Top (Y) position in points, sorts the shapes from top to bottom, and prints their Name, Top and Left values. | Extend the shape‑sorting example to additionally order shapes by their Left (X) coordinate after the Y‑axis sort and write the sorted list to a new Excel file.
// Common Searches: aspocells get absolute Y coordinate of shapes in worksheet c# | sort excel worksheet shapes by vertical position using Aspose.Cells | list shape names and positions in an Excel file with Aspose.Cells .NET | how to order shapes by top property in Aspose.Cells | retrieve shape top and left values from workbook using C#
// Tags: Aspose.Cells enumerate worksheet shapes | Aspose.Cells retrieve shape top property | Aspose.Cells sort shapes by vertical coordinate | Aspose.Cells output shape name and position | Aspose.Cells save workbook after shape ordering

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// The program loads an Excel workbook, gathers each shape on the first worksheet with its absolute Top (Y) and Left (X) coordinates, sorts the shapes by the Y coordinate, prints the sorted order (name, top, left), and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Collect each shape together with its absolute Y coordinate (Top)
            var shapeData = new List<(Shape shape, double top)>();

            foreach (Shape shape in worksheet.Shapes)
            {
                // Top property gives the absolute Y position in points
                double topPosition = shape.Top;
                shapeData.Add((shape, topPosition));
            }

            // Sort the shapes by their Y coordinate (ascending)
            var sortedShapes = shapeData.OrderBy(item => item.top).ToList();

            // Output the sorted order
            Console.WriteLine("Shapes sorted by Y coordinate (Top):");
            for (int i = 0; i < sortedShapes.Count; i++)
            {
                Shape s = sortedShapes[i].shape;
                Console.WriteLine($"{i + 1}. Name: {s.Name}, Top: {s.Top}, Left: {s.Left}");
            }

            // Save the workbook if any changes were made (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
