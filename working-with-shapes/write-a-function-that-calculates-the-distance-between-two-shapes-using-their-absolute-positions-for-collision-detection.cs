using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDistanceApp
{
    public static class ShapeCollisionHelper
    {
        // Calculates Euclidean distance between the centers of two shapes using absolute positions.
        public static double GetDistance(Shape shape1, Shape shape2)
        {
            double centerX1 = (shape1.Left + shape1.Right) / 2.0;
            double centerY1 = (shape1.Top + shape1.Bottom) / 2.0;

            double centerX2 = (shape2.Left + shape2.Right) / 2.0;
            double centerY2 = (shape2.Top + shape2.Bottom) / 2.0;

            double dx = centerX2 - centerX1;
            double dy = centerY2 - centerY1;

            return Math.Sqrt(dx * dx + dy * dy);
        }
    }

    public class ShapeDistanceDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle start)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add two rectangle shapes with distinct positions
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
                Shape shapeA = sheet.Shapes.AddRectangle(2, 0, 50, 50, 100, 100);
                Shape shapeB = sheet.Shapes.AddRectangle(5, 0, 200, 300, 120, 80);

                // Compute distance between the two shapes
                double distance = ShapeCollisionHelper.GetDistance(shapeA, shapeB);
                Console.WriteLine($"Distance between Shape A and Shape B: {distance} pixels");

                // Save the workbook (lifecycle end)
                string outputPath = "ShapeDistanceDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeDistanceDemo.Run();
        }
    }
}