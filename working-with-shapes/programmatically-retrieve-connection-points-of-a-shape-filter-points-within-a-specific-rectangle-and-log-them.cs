using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConnectionPointsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

            // Retrieve all connection points of the shape
            // Each point is a float[2] where [0] = X and [1] = Y (in points)
            float[][] connectionPoints = shape.GetConnectionPoints();

            // Define the rectangle area for filtering (example values)
            // These values are in the same coordinate system as the connection points
            float rectLeft = 20f;   // minimum X
            float rectTop = 20f;    // minimum Y
            float rectRight = 80f;  // maximum X
            float rectBottom = 80f; // maximum Y

            Console.WriteLine("Connection points within the rectangle:");
            for (int i = 0; i < connectionPoints.Length; i++)
            {
                float x = connectionPoints[i][0];
                float y = connectionPoints[i][1];

                // Check if the point lies inside the defined rectangle
                if (x >= rectLeft && x <= rectRight && y >= rectTop && y <= rectBottom)
                {
                    Console.WriteLine($"Point {i + 1}: X={x}, Y={y}");
                }
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("ConnectionPointsDemo.xlsx");
        }
    }
}