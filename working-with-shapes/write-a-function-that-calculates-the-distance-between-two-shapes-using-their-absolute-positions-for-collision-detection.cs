// Title: C# – Compute Euclidean distance between two Aspose.Cells shapes using absolute coordinates
// Description: Demonstrates how to obtain the center points of two Aspose.Cells Shape objects via their Left, Right, Top and Bottom properties, calculate the Euclidean distance between those centers, and use the result for collision‑detection or layout logic in a .NET workbook.
// Keywords: Aspose.Cells shape distance C# | calculate shape center Aspose.Cells | collision detection worksheet shapes .NET | Euclidean distance between shapes Aspose | absolute position shape Aspose.Cells
// Common Searches: how to measure distance between two shapes in Aspose.Cells | C# get shape center coordinates Aspose.Cells | collision detection for worksheet shapes .NET | determine spacing of shapes in an Excel file using Aspose | calculate Euclidean distance of Aspose.Cells shapes
// Developer Intent: Find a reliable way to measure the straight‑line distance between two worksheet shapes for proximity checks or layout adjustments.
// Use Cases: Validate minimum clearance between diagram elements in an automatically generated report. | Trigger automatic repositioning when shapes become too close during dynamic chart creation. | Generate proximity metrics for custom layout engines that arrange shapes based on relative distance.
// AI Prompts: Write a C# method that returns true if two Aspose.Cells shapes intersect using their absolute positions. | Create an extension method to compute Manhattan distance between Aspose.Cells shapes. | Produce unit tests for the distance function covering overlapping, adjacent, and distant shape scenarios.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to obtain the center points of two Aspose.Cells Shape objects via their Left, Right, Top and Bottom properties, calculate the Euclidean distance between those centers, and use the result for collision‑detection or layout logic in a .NET workbook.
public class ShapeDistanceCalculator
{
    // Calculates Euclidean distance between the centers of two shapes
    public static double GetDistanceBetweenShapes(Shape shape1, Shape shape2)
    {
        // Center X = Left + (Right - Left) / 2
        double x1 = shape1.Left + (shape1.Right - shape1.Left) / 2.0;
        double y1 = shape1.Top + (shape1.Bottom - shape1.Top) / 2.0;

        double x2 = shape2.Left + (shape2.Right - shape2.Left) / 2.0;
        double y2 = shape2.Top + (shape2.Bottom - shape2.Top) / 2.0;

        double dx = x2 - x1;
        double dy = y2 - y1;

        return Math.Sqrt(dx * dx + dy * dy);
    }

    // Demonstrates usage of the distance calculation
    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add first rectangle shape and define its position and size
        Shape shapeA = sheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);
        shapeA.Left = 50;               // X position
        shapeA.Top = 50;                // Y position
        shapeA.Right = shapeA.Left + 100;   // Width = 100
        shapeA.Bottom = shapeA.Top + 100;   // Height = 100

        // Add second rectangle shape and define its position and size
        Shape shapeB = sheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);
        shapeB.Left = 200;
        shapeB.Top = 150;
        shapeB.Right = shapeB.Left + 100;
        shapeB.Bottom = shapeB.Top + 100;

        // Calculate and output the distance between the two shapes
        double distance = GetDistanceBetweenShapes(shapeA, shapeB);
        Console.WriteLine($"Distance between shapes: {distance}");

        // Save the workbook (optional, demonstrates lifecycle handling)
        workbook.Save("ShapeDistanceDemo.xlsx");
    }

    // Entry point required for console application
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
