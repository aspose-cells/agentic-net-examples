// Title: Aspose.Cells .NET – Retrieve shape connection points and list those on the right edge
// Description: Creates a workbook, adds a rectangle shape, calculates its right‑edge X coordinate, reads all connection points with GetConnectionPoints, filters points that lie on the right edge using a tolerance, logs the matching points, and saves the file.
// Keywords: Aspose.Cells shape connection points | C# GetConnectionPoints | filter right edge points | shape bounding rectangle Aspose.Cells | floating point tolerance comparison | Aspose.Cells connector alignment
// Common Searches: Aspose.Cells get shape connection points .NET | filter shape connection points by right side | C# Aspose.Cells right edge connection points | how to read shape connectors in Aspose.Cells | Aspose.Cells shape GetConnectionPoints example
// Developer Intent: Read all connection points of a worksheet shape and output only those that are positioned on the shape’s right edge.
// Use Cases: Attach connectors programmatically to the right side of a shape using filtered connection points. | Validate layout consistency by ensuring custom connectors align with the shape’s right edge before saving. | Generate a coordinate report of right‑edge connection points for documentation or auditing.
// AI Prompts: Generate C# code with Aspose.Cells that reads a shape’s connection points and returns those on the right edge within a tolerance. | Show how to compute a shape’s right‑edge X coordinate and compare it to connection point X values in Aspose.Cells .NET. | Provide an example that logs right‑edge connection points of a rectangle shape and saves the workbook using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, calculates its right‑edge X coordinate, reads all connection points with GetConnectionPoints, filters points that lie on the right edge using a tolerance, logs the matching points, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Build the bounding rectangle of the shape using its position and size properties
            RectangleF rect = new RectangleF(shape.Left, shape.Top, shape.Width, shape.Height);

            // X coordinate of the right edge of the shape
            float rightEdgeX = rect.X + rect.Width;

            // Retrieve all connection points of the shape
            float[][] connectionPoints = shape.GetConnectionPoints();

            Console.WriteLine("Connection points located on the right edge of the shape:");
            const float tolerance = 0.01f; // tolerance for floating‑point comparison

            for (int i = 0; i < connectionPoints.Length; i++)
            {
                float x = connectionPoints[i][0];
                float y = connectionPoints[i][1];

                // Identify points whose X coordinate matches the right edge (within tolerance)
                if (Math.Abs(x - rightEdgeX) <= tolerance)
                {
                    Console.WriteLine($"Point {i + 1}: X = {x}, Y = {y}");
                }
            }

            // Define output file path
            string outputPath = "ConnectionPointsDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
