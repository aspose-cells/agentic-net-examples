using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeConnectionPointsDemo
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (parameters: upper left row, column, top, left, height, width, shape type)
            // Here we use AddRectangle which returns a Shape object
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Retrieve all connection points of the shape
            // Each point is a float[2] where [0] = X and [1] = Y
            float[][] connectionPoints = shape.GetConnectionPoints();

            // Filter points where X coordinate is greater than 100
            Console.WriteLine("Connection points with X > 100:");
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                float x = connectionPoints[i][0];
                float y = connectionPoints[i][1];

                if (x > 100f)
                {
                    Console.WriteLine($"Point {i + 1}: X = {x}, Y = {y}");
                }
            }

            // Save the workbook (optional, just to demonstrate lifecycle handling)
            workbook.Save("ShapeConnectionPointsDemo.xlsx");
        }
    }
}