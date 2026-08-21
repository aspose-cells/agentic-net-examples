// Title: C# Example: Retrieve and Filter Shape Connection Points (X > 100) Using Aspose.Cells
// Description: This Aspose.Cells for .NET sample creates a workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to obtain all connection points, iterates the float[][] array, logs only points whose X coordinate exceeds 100, and saves the file.
// Keywords: Aspose.Cells C# GetConnectionPoints | shape connection points Aspose.Cells | filter shape points by X coordinate | Aspose.Cells rectangle shape example | C# retrieve shape connection points | Aspose.Cells connection points filtering | Shape.GetConnectionPoints C#
// Common Searches: Aspose.Cells get connection points C# | filter shape points X > 100 Aspose.Cells | Shape.GetConnectionPoints example | C# retrieve rectangle shape connection points | Aspose.Cells log connection points | how to filter shape connection points Aspose.Cells
// Developer Intent: The developer wants to obtain a shape’s connection points, keep only those with an X value greater than 100, and output them.
// Use Cases: Identify anchor points for custom connectors when the X coordinate exceeds a threshold. | Generate a report of shape connection points for layout analysis or validation. | Programmatically align additional shapes based on filtered connection points. | Validate shape placement against design rules that depend on X‑coordinate limits. | Export the filtered points to another system for further processing.
// AI Prompts: Provide C# code using Aspose.Cells to retrieve all connection points of a shape and filter them where X > 100. | Explain the structure of the float[][] returned by Shape.GetConnectionPoints and how to access X and Y values. | Show how to modify the example to filter by Y coordinate or to store the filtered points in a collection for later use.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This Aspose.Cells for .NET sample creates a workbook, adds a rectangle shape, calls Shape.GetConnectionPoints() to obtain all connection points, iterates the float[][] array, logs only points whose X coordinate exceeds 100, and saves the file.
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

        // Retrieve all connection points of the shape
        float[][] connectionPoints = shape.GetConnectionPoints();

        // Log points whose X coordinate is greater than 100
        Console.WriteLine("Connection points with X > 100:");
        for (int i = 0; i < connectionPoints.Length; i++)
        {
            float x = connectionPoints[i][0];
            float y = connectionPoints[i][1];

            if (x > 100)
            {
                Console.WriteLine($"Point {i + 1}: X = {x}, Y = {y}");
            }
        }

        // Save the workbook (optional, demonstrates lifecycle handling)
        workbook.Save("ConnectionPointsDemo.xlsx");
    }
}
