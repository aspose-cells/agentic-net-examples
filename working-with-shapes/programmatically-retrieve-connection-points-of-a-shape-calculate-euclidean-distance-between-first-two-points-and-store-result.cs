// Title: C# – Retrieve shape connection points and calculate Euclidean distance using Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), computes the Euclidean distance between the first two points, writes the result to cell A1, and saves the file as ShapeConnectionPointsDistance.xlsx.
// Keywords: Aspose.Cells GetConnectionPoints | C# shape connection points | calculate Euclidean distance in Excel | store geometry result in worksheet | Aspose.Cells rectangle shape example
// Common Searches: Aspose.Cells get shape connection points C# | how to compute distance between shape points in Aspose.Cells | write calculated distance to Excel cell using Aspose.Cells | example of GetConnectionPoints with rectangle shape
// Developer Intent: Extract a shape's connection points, measure the distance between the first two points, and record the value in a worksheet cell.
// Use Cases: Validate layout dimensions by measuring distances between key shape anchors. | Generate reports that include geometric metrics directly in Excel. | Automate shape‑based calculations for engineering or design spreadsheets.
// AI Prompts: Generate C# code to loop through all connection points of a shape and output pairwise distances. | Show how to convert the float[][] returned by GetConnectionPoints into System.Drawing.PointF objects. | Explain error handling when a shape provides fewer than two connection points in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), computes the Euclidean distance between the first two points, writes the result to cell A1, and saves the file as ShapeConnectionPointsDistance.xlsx.
class Program
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
        float[][] points = shape.GetConnectionPoints();

        // Calculate Euclidean distance between the first two points, if they exist
        if (points != null && points.Length >= 2)
        {
            float x1 = points[0][0];
            float y1 = points[0][1];
            float x2 = points[1][0];
            float y2 = points[1][1];

            double distance = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));

            // Store the calculated distance in cell A1
            worksheet.Cells["A1"].PutValue(distance);
        }

        // Save the workbook with the result
        workbook.Save("ShapeConnectionPointsDistance.xlsx");
    }
}
