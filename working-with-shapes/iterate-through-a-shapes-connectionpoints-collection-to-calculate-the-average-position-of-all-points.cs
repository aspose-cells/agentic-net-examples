// Title: Compute the centroid of a shape’s connection points with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle shape, retrieves its connection points via GetConnectionPoints(), safely handles null or empty collections, sums X and Y values, calculates the average coordinates, prints the centroid, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | shape connection points | average coordinates | centroid | GetConnectionPoints | Excel shape geometry
// Common Searches: Aspose.Cells calculate shape centroid | C# average connection point Aspose.Cells | GetConnectionPoints average X Y | find geometric center of Excel shape using Aspose.Cells | iterate shape connection points C#
// Developer Intent: Determine a shape’s geometric center by averaging its connection point coordinates.
// Use Cases: Place a label or comment at the exact center of a custom shape. | Align multiple objects relative to the centroid of a reference shape. | Validate symmetry of a shape’s connection points during automated diagram generation.
// AI Prompts: Write C# code that uses Aspose.Cells to retrieve a shape’s connection points and compute the average X and Y values, including null‑check handling. | Explain how to use the centroid of a shape to position another worksheet element with Aspose.Cells. | Provide a step‑by‑step tutorial for iterating over Shape.GetConnectionPoints() and calculating the centroid, covering error handling for missing or incomplete point data.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle shape, retrieves its connection points via GetConnectionPoints(), safely handles null or empty collections, sums X and Y values, calculates the average coordinates, prints the centroid, and saves the file.
    class ShapeConnectionPointsAverage
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width, shape type (0 = rectangle)
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Retrieve the connection points of the shape
            // Each point is a float[2] array where [0] = X and [1] = Y
            float[][] points = shape.GetConnectionPoints();

            // Guard against shapes with no connection points
            if (points == null || points.Length == 0)
            {
                Console.WriteLine("The shape has no connection points.");
            }
            else
            {
                // Sum X and Y coordinates
                double sumX = 0;
                double sumY = 0;
                foreach (float[] pt in points)
                {
                    // Ensure each point has both coordinates
                    if (pt != null && pt.Length >= 2)
                    {
                        sumX += pt[0];
                        sumY += pt[1];
                    }
                }

                // Calculate averages
                double avgX = sumX / points.Length;
                double avgY = sumY / points.Length;

                // Output the result
                Console.WriteLine($"Average Connection Point: X = {avgX}, Y = {avgY}");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ShapeConnectionPointsAverage.xlsx");
        }
    }
}
