// Title: C# – Compute pixel distance between Aspose.Cells shapes using absolute positions for collision detection
// Description: Demonstrates how to use Aspose.Cells' Shape.GetActualBox() in C# to obtain each shape's bounding box, calculate the center points, and return the Euclidean distance between two worksheet shapes. The sample creates two rectangles, sets their Left/Top coordinates, prints the pixel distance, and saves the workbook.
// Keywords: Aspose.Cells | C# | shape distance | GetActualBox | collision detection | pixel distance | worksheet shapes | Euclidean distance | shape positioning | Aspose.Cells API
// Common Searches: Aspose.Cells calculate distance between shapes C# | GetActualBox shape collision detection | pixel distance worksheet shapes Aspose | C# compute Euclidean distance for Aspose.Cells shapes | how to measure shape separation in Excel using Aspose
// Developer Intent: Obtain the pixel distance between two worksheet shapes to enable collision detection or layout adjustments.
// Use Cases: Check whether two shapes overlap and reposition them to avoid collision. | Dynamically space shapes based on a minimum distance threshold. | Draw connector lines whose length matches the computed distance between objects.
// AI Prompts: Generate a method that returns true when two Aspose.Cells shapes are within a given pixel range using GetActualBox. | Extend the distance function to support ovals, lines, and other shape types, returning the shortest distance between perimeters. | Create unit tests for GetDistance covering identical positions, far‑apart shapes, and partially overlapping cases.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDistance
{
    // Demonstrates how to use Aspose.Cells' Shape.GetActualBox() in C# to obtain each shape's bounding box, calculate the center points, and return the Euclidean distance between two worksheet shapes. The sample creates two rectangles, sets their Left/Top coordinates, prints the pixel distance, and saves the workbook.
    public class ShapeDistanceCalculator
    {
        // Calculates Euclidean distance between the centers of two shapes
        public static double GetDistance(Shape shape1, Shape shape2)
        {
            // Get actual bounding boxes: [x, y, width, height]
            float[] box1 = shape1.GetActualBox();
            float[] box2 = shape2.GetActualBox();

            // Compute center points
            double centerX1 = box1[0] + box1[2] / 2.0;
            double centerY1 = box1[1] + box1[3] / 2.0;

            double centerX2 = box2[0] + box2[2] / 2.0;
            double centerY2 = box2[1] + box2[3] / 2.0;

            // Euclidean distance
            double deltaX = centerX2 - centerX1;
            double deltaY = centerY2 - centerY1;
            return Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
        }

        // Example usage
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two rectangle shapes with different positions
            Shape rect1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 80); // Left=2, Top=0, Width=100, Height=80
            Shape rect2 = worksheet.Shapes.AddRectangle(5, 0, 200, 150, 120, 90); // Left=5, Top=0, Width=120, Height=90

            // Optionally adjust positions using properties
            rect1.Left = 50;
            rect1.Top = 30;

            rect2.Left = 300;
            rect2.Top = 200;

            // Calculate distance between the two shapes
            double distance = GetDistance(rect1, rect2);
            Console.WriteLine($"Distance between shapes: {distance:F2} pixels");

            // Save the workbook (optional, just to demonstrate lifecycle)
            workbook.Save("ShapeDistanceDemo.xlsx");
        }
    }

    class Program
    {
        static void Main()
        {
            ShapeDistanceCalculator.Run();
        }
    }
}
