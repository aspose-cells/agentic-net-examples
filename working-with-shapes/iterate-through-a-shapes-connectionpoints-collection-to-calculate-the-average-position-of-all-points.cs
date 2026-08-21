// Title: C# – Compute the Average Position of a Shape’s Connection Points with Aspose.Cells
// Description: This example creates a workbook, adds a rectangle shape, retrieves its connection points using Shape.GetConnectionPoints(), checks for null or empty collections, sums the X and Y values, calculates the average coordinates, prints the centroid, and saves the file.
// Keywords: Aspose.Cells C# | shape connection points | GetConnectionPoints | average coordinates | centroid calculation | iterate connection points | worksheet shape geometry | null check shape points
// Common Searches: how to get average connection point of a shape in Aspose.Cells | C# iterate Shape.GetConnectionPoints() Aspose.Cells | calculate centroid of rectangle shape using Aspose.Cells | Aspose.Cells shape connection points null handling | average X Y of shape points .NET
// Developer Intent: Find the mean X‑Y location of all connection points belonging to a shape.
// Use Cases: Determine the geometric center of a custom shape to align other objects. | Place a marker or annotation at the computed centroid for visual reference. | Validate shape layout by comparing the calculated average point with design specifications.
// AI Prompts: Generate a reusable C# method that returns the centroid of any shape’s connection points using Aspose.Cells. | Show how to safely handle shapes with no connection points before performing average calculations. | Provide code that adds a small marker shape at the calculated average connection point on the worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a rectangle shape, retrieves its connection points using Shape.GetConnectionPoints(), checks for null or empty collections, sums the X and Y values, calculates the average coordinates, prints the centroid, and saves the file.
class ShapeConnectionPointsAverage
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width, shape type (0 = rectangle)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Retrieve the connection points of the shape
        // Each point is a float[2] where [0] = X and [1] = Y
        float[][] connectionPoints = shape.GetConnectionPoints();

        // Guard against shapes with no connection points
        if (connectionPoints == null || connectionPoints.Length == 0)
        {
            Console.WriteLine("The shape has no connection points.");
        }
        else
        {
            // Calculate the sum of X and Y coordinates
            float sumX = 0f;
            float sumY = 0f;
            foreach (float[] point in connectionPoints)
            {
                // Ensure each point has both X and Y values
                if (point != null && point.Length >= 2)
                {
                    sumX += point[0];
                    sumY += point[1];
                }
            }

            // Compute the average position
            float avgX = sumX / connectionPoints.Length;
            float avgY = sumY / connectionPoints.Length;

            Console.WriteLine($"Average Connection Point: X = {avgX}, Y = {avgY}");
        }

        // Save the workbook (optional, just to demonstrate lifecycle compliance)
        workbook.Save("ShapeConnectionPointsAverage.xlsx");
    }
}
