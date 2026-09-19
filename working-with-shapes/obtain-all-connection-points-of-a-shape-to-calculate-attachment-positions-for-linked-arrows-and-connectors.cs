// Title: Retrieve all connection points of a named shape and compute their absolute coordinates in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Extract the collection of connection points from a shape named "MyShape" and compute each point's absolute X/Y coordinates in points using the shape's width and height with Aspose.Cells in C#. | Generate C# code that safely checks for a missing shape or an empty connection‑point collection before iterating through GetConnectionPoints. | Create a method that returns a list of objects containing the index, relative X/Y (0‑1) and absolute X/Y (points) for every connection point of any worksheet shape.
// Common Searches: Aspose.Cells C# get connection points of a drawing shape in Excel | calculate absolute position of shape connection points using shape dimensions Aspose.Cells | how to handle GetConnectionPoints returning null in Aspose.Cells .NET | retrieve index and relative coordinates of shape connectors in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells GetConnectionPoints API | C# calculate absolute shape connector coordinates | Excel shape connection point extraction | handling missing shape Aspose.Cells | dynamic shape connection points iteration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file, accesses the first worksheet, finds a shape named "MyShape", obtains its connection points via GetConnectionPoints, and for each point calculates absolute X/Y positions (in points) using the shape's width and height, then outputs the index, relative coordinates, and absolute coordinates.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the shape by its name
            Shape shape = worksheet.Shapes["MyShape"];
            if (shape == null)
            {
                Console.WriteLine("Shape not found.");
                return;
            }

            // Get the collection of connection points (if any) using dynamic to avoid compile‑time type dependencies
            dynamic cpCollection = shape.GetConnectionPoints();
            if (cpCollection == null || cpCollection.Count == 0)
            {
                Console.WriteLine("No connection points found for the shape.");
                return;
            }

            // Iterate through all connection points of the shape
            foreach (dynamic cp in cpCollection)
            {
                // Index of the connection point
                int index = cp.Index;

                // Relative positions (0.0 – 1.0) within the shape
                double relativeX = cp.X;
                double relativeY = cp.Y;

                // Calculate absolute position in points using shape dimensions
                double absoluteX = shape.Width * relativeX;
                double absoluteY = shape.Height * relativeY;

                Console.WriteLine($"Connection Point {index}:");
                Console.WriteLine($"  Relative Position -> X: {relativeX:F2}, Y: {relativeY:F2}");
                Console.WriteLine($"  Absolute Position (points) -> X: {absoluteX:F2}, Y: {absoluteY:F2}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
