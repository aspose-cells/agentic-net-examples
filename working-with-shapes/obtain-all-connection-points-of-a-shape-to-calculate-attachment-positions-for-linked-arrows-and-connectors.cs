using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Get all connection points of the shape
        float[][] points = shape.GetConnectionPoints();

        // Output the connection points (X,Y) pairs
        Console.WriteLine("Connection Points:");
        for (int i = 0; i < points.Length; i++)
        {
            Console.WriteLine($"Point {i + 1}: X={points[i][0]}, Y={points[i][1]}");
        }

        // Optional: save the workbook if needed
        // workbook.Save("ConnectionPointsDemo.xlsx");
    }
}