// Title: C# – Get and filter shape connection points (X > 100) with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a rectangle shape, retrieve its connection points via Shape.GetConnectionPoints(), filter points whose X coordinate exceeds 100, log the matching points, and save the file.
// Keywords: Aspose.Cells | C# | Shape.GetConnectionPoints | connection points | filter by X coordinate | rectangle shape | Aspose.Cells .NET example | retrieve shape anchors | Excel shape connection points | Aspose.Cells shape API
// Common Searches: Aspose.Cells get shape connection points C# | filter shape connection points where X > 100 | Shape.GetConnectionPoints example .NET | how to list connection points of a rectangle in Aspose.Cells | retrieve shape anchors Aspose.Cells | C# code to log shape connection points | save workbook after processing shapes Aspose.Cells
// Developer Intent: Extract a shape’s connection points, keep only those with X > 100, and output them (e.g., to console or log).
// Use Cases: Identify anchor locations for custom connectors based on horizontal position. | Validate shape layout by checking that connection points do not cross a defined X‑coordinate threshold. | Generate a diagnostic report of connection points before exporting the worksheet. | Programmatically adjust connectors or align shapes using filtered connection points.
// AI Prompts: Write a C# snippet using Aspose.Cells that adds a rectangle shape, calls Shape.GetConnectionPoints(), filters points with X > 100, prints them, and saves the workbook. | Explain the structure of the float[][] returned by Shape.GetConnectionPoints and demonstrate how to iterate and filter by X coordinate in .NET. | Provide a step‑by‑step guide to retrieve and log shape connection points in Aspose.Cells, including error handling and workbook saving.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConnectionPointsDemo
{
    // Shows how to create a workbook, add a rectangle shape, retrieve its connection points via Shape.GetConnectionPoints(), filter points whose X coordinate exceeds 100, log the matching points, and save the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
                Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 200);

                // Retrieve all connection points of the shape
                float[][] connectionPoints = shape.GetConnectionPoints();

                // Log connection points where the X coordinate is greater than 100
                Console.WriteLine("Connection points with X > 100:");
                for (int i = 0; i < connectionPoints.Length; i++)
                {
                    float x = connectionPoints[i][0];
                    float y = connectionPoints[i][1];

                    if (x > 100)
                    {
                        Console.WriteLine($"Point {i + 1}: X = {x}, Y = {y}");
                    }
                }

                // Save the workbook (optional, just to demonstrate lifecycle handling)
                string outputPath = "ConnectionPointsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
