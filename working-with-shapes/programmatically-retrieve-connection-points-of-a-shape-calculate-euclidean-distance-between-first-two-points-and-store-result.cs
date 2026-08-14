// Title: Aspose.Cells for .NET: Retrieve Shape Connection Points and Compute Euclidean Distance
// Description: This C# example creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), verifies at least two points, calculates the Euclidean distance between the first two points, writes the result (or an error note) to cell C1, and saves the file as ShapeConnectionPointsDistance.xlsx.
// Keywords: Aspose.Cells | .NET | C# shape connection points | GetConnectionPoints | Euclidean distance calculation | Excel geometry measurement | write value to cell | save workbook | shape analysis Aspose
// Common Searches: Aspose.Cells get shape connection points C# | calculate distance between two shape points .NET | write computed value to Excel cell using Aspose | how to measure shape geometry with Aspose.Cells
// Developer Intent: Extract a shape’s connection points, compute the distance between the first two points, and store the numeric result in a worksheet cell.
// Use Cases: Validate that a drawn shape meets design specifications by measuring point-to-point distance. | Include geometric metrics of shapes directly in generated Excel reports. | Detect shapes lacking sufficient connection points and log a descriptive message in the sheet.
// AI Prompts: Show how to loop through all consecutive connection point pairs and output each distance to separate cells. | Provide code that rounds the calculated distance to two decimal places before writing it to the worksheet. | Explain how to select the farthest pair of connection points when a shape returns more than two points and compute that distance.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), verifies at least two points, calculates the Euclidean distance between the first two points, writes the result (or an error note) to cell C1, and saves the file as ShapeConnectionPointsDistance.xlsx.
    public class ShapeConnectionPointsDistance
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

            // Retrieve the connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            // Ensure there are at least two connection points
            if (points != null && points.Length >= 2 && points[0].Length >= 2 && points[1].Length >= 2)
            {
                // First point (X1, Y1)
                float x1 = points[0][0];
                float y1 = points[0][1];

                // Second point (X2, Y2)
                float x2 = points[1][0];
                float y2 = points[1][1];

                // Calculate Euclidean distance between the first two points
                double distance = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));

                // Store the distance value in cell C1
                worksheet.Cells["C1"].PutValue(distance);
            }
            else
            {
                // If insufficient points, store a message
                worksheet.Cells["C1"].PutValue("Insufficient connection points");
            }

            // Save the workbook to a file
            string outputPath = "ShapeConnectionPointsDistance.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
