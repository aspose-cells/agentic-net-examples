// Title: C# Aspose.Cells – Retrieve and Filter Shape Connection Points Within a Rectangle
// Description: Demonstrates adding a rectangle shape to a workbook, using GetConnectionPoints to obtain its connection points, filtering those that lie inside a defined rectangle, and logging the matching coordinates.
// Keywords: Aspose.Cells | C# | .NET | GetConnectionPoints | shape connection points | rectangle filter | Aspose.Cells.Drawing | code sample | workbook shape example | Aspose.Cells API
// Common Searches: Aspose.Cells get shape connection points C# | filter shape points inside rectangle Aspose.Cells | C# example GetConnectionPoints Aspose.Cells | how to retrieve shape connection points in .NET | Aspose.Cells shape connector coordinates | sample code for shape point filtering Aspose.Cells
// Developer Intent: Retrieve a shape’s connection points, keep only those that fall within a specific rectangular area, and output the filtered list.
// Use Cases: Identify connection points that are inside a printable region before attaching connectors. | Validate that shape connection points stay within custom layout boundaries to prevent overlap. | Extract spatially‑qualified points for further processing in reporting or diagram generation.
// AI Prompts: Generate C# code using Aspose.Cells to list all connection points of a shape and return only those inside a given rectangle. | Explain the structure of the float[][] array returned by GetConnectionPoints and how to interpret X/Y values. | Show how to modify the example to filter connection points by a circular radius instead of a rectangle.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates adding a rectangle shape to a workbook, using GetConnectionPoints to obtain its connection points, filtering those that lie inside a defined rectangle, and logging the matching coordinates.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Retrieve all connection points of the shape
        float[][] connectionPoints = shape.GetConnectionPoints();

        // Define the rectangle area for filtering (example values)
        float filterLeft   = 20f;   // minimum X
        float filterTop    = 30f;   // minimum Y
        float filterRight  = 80f;   // maximum X
        float filterBottom = 150f;  // maximum Y

        // Log points that lie within the defined rectangle
        Console.WriteLine("Connection points inside the rectangle:");
        for (int i = 0; i < connectionPoints.Length; i++)
        {
            float x = connectionPoints[i][0];
            float y = connectionPoints[i][1];

            if (x >= filterLeft && x <= filterRight && y >= filterTop && y <= filterBottom)
            {
                Console.WriteLine($"Point {i + 1}: X = {x}, Y = {y}");
            }
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ConnectionPointsDemo.xlsx");
    }
}
