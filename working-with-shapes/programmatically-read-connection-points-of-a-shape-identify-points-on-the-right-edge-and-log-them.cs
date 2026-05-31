using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to work with
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Retrieve all connection points of the shape
        float[][] points = shape.GetConnectionPoints();

        // Find the maximum X coordinate (rightmost edge)
        float maxX = float.MinValue;
        foreach (float[] pt in points)
        {
            if (pt[0] > maxX)
                maxX = pt[0];
        }

        // Define a tolerance for floating‑point comparison
        const float tolerance = 0.01f;

        // Log points that lie on the right edge
        Console.WriteLine("Connection points on the right edge:");
        for (int i = 0; i < points.Length; i++)
        {
            if (Math.Abs(points[i][0] - maxX) <= tolerance)
            {
                Console.WriteLine($"Point {i + 1}: X={points[i][0]}, Y={points[i][1]}");
            }
        }

        // Save the workbook (save rule)
        workbook.Save("ConnectionPointsDemo.xlsx");
    }
}