using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

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
        // Each point is a float[2] array where [0] = X and [1] = Y
        float[][] connectionPoints = shape.GetConnectionPoints();

        // Calculate the average X and Y coordinates
        double sumX = 0;
        double sumY = 0;
        int pointCount = connectionPoints.Length;

        for (int i = 0; i < pointCount; i++)
        {
            sumX += connectionPoints[i][0];
            sumY += connectionPoints[i][1];
        }

        double avgX = pointCount > 0 ? sumX / pointCount : 0;
        double avgY = pointCount > 0 ? sumY / pointCount : 0;

        Console.WriteLine($"Average Connection Point: X = {avgX}, Y = {avgY}");

        // Save the workbook to a file
        workbook.Save("ShapeConnectionPointsAverage.xlsx");
    }
}