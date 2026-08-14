// Title: C# – Retrieve shape connection points and log those on the right edge with Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, reads its connection points via GetConnectionPoints(), calculates the shape's right‑edge X coordinate using the Width property, and writes to the console only the points that lie on that edge within a tolerance.
// Keywords: Aspose.Cells shape connection points | GetConnectionPoints C# | filter connection points by X coordinate | right edge shape Aspose.Cells | shape width tolerance Aspose.Cells | log shape connection points .NET
// Common Searches: Aspose.Cells get shape connection points C# | how to find right‑edge connection points in Aspose.Cells | filter shape connection points by X value Aspose.Cells | calculate shape right edge using Width property Aspose.Cells | C# example for reading shape connection points
// Developer Intent: Read a shape's connection points and output only those positioned on the shape's right side.
// Use Cases: Verify that connectors are attached to the right side of a diagram element before exporting. | Produce a layout report that lists all right‑edge connection points for auditing Excel drawings. | Programmatically adjust connector positions based on right‑edge connection points of shapes.
// AI Prompts: Generate C# code that reads shape connection points with Aspose.Cells and filters those on the right edge using a tolerance. | Explain how to compute the right‑edge X coordinate of a shape and compare it to connection point coordinates in Aspose.Cells. | Show an example of logging only right‑edge connection points of a rectangle shape in a .NET console application.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle shape, reads its connection points via GetConnectionPoints(), calculates the shape's right‑edge X coordinate using the Width property, and writes to the console only the points that lie on that edge within a tolerance.
    public class ShapeConnectionPointsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, row offset, column offset, height, width
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

                // Retrieve the connection points of the shape
                float[][] points = shape.GetConnectionPoints();

                // Determine the right edge X coordinate of the shape
                double rightEdgeX = shape.Width; // Width is in points

                // Tolerance for floating point comparison
                const double tolerance = 0.5;

                // Log all connection points that lie on the right edge
                Console.WriteLine("Connection points on the right edge:");
                for (int i = 0; i < points.Length; i++)
                {
                    double x = points[i][0];
                    double y = points[i][1];

                    if (Math.Abs(x - rightEdgeX) <= tolerance)
                    {
                        Console.WriteLine($"Point {i + 1}: X={x}, Y={y}");
                    }
                }

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                workbook.Save("ShapeConnectionPointsDemo.xlsx");
                Console.WriteLine("Workbook saved as ShapeConnectionPointsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeConnectionPointsDemo.Run();
        }
    }
}
