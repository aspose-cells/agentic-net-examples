using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeConnectionPointsDistanceDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (row, column, row offset, column offset, height, width)
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

            // Retrieve all connection points of the shape
            float[][] points = shape.GetConnectionPoints();

            // Ensure there are at least two valid connection points
            if (points != null && points.Length >= 2 &&
                points[0] != null && points[0].Length >= 2 &&
                points[1] != null && points[1].Length >= 2)
            {
                float x1 = points[0][0];
                float y1 = points[0][1];
                float x2 = points[1][0];
                float y2 = points[1][1];

                // Calculate Euclidean distance between the two points
                double distance = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));

                // Store the distance value in cell B2
                worksheet.Cells["B2"].PutValue(distance);
            }
            else
            {
                // Insufficient points – store a sentinel message
                worksheet.Cells["B2"].PutValue("Insufficient connection points");
            }

            string outputPath = "ShapeConnectionPointsDistanceDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}